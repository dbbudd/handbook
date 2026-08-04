import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { mark, scheduleDump } from '../shared/hbTiming';

const LOG_SOURCE = 'GlossaryTooltipsApplicationCustomizer';
const HANDBOOK_PAGE_PATH = '/sitepages/home1.aspx';

// Module-evaluation mark — see hbTiming.ts. This is a SEPARATE bundle from
// HandbookExperience, so comparing the two "bundle evaluated" marks also shows
// whether SharePoint loads the two customizers together or staggers them.
mark('glossary', 'bundle evaluated');

const stripTrailingSlash = (s: string): string => s.replace(/\/+$/, '');

/**
 * Same widened guard as HandbookExperience — see the long comment there. Kept as
 * a local copy rather than shared because the two customizers are deliberately
 * independent (CLAUDE.md §3: each must survive the other failing), and a shared
 * guard would couple their activation logic.
 */
const matchHandbookPath = (
  pathname: string,
  targetPath: string,
  webServerRelativeUrl: string
): 'page' | 'site-root' | undefined => {
  const path = stripTrailingSlash(pathname.toLowerCase());
  if (path.indexOf(targetPath, path.length - targetPath.length) !== -1) return 'page';
  const web = stripTrailingSlash((webServerRelativeUrl || '').toLowerCase());
  if (web.length > 0 && path === web) return 'site-root';
  return undefined;
};
const POPOVER_ID = 'hb-term-popover';

const LIST_TITLE = 'Glossary';
const CACHE_KEY = 'hb_glossary_v1';
const CACHE_TTL_MS = 10 * 60 * 1000; // 10 minutes
const MAX_TERMS = 500;

interface IGlossaryTerm {
  term: string;
  definition: string;
}

export interface IGlossaryTooltipsApplicationCustomizerProperties {
  pagePath?: string;
  listTitle?: string;
}

export default class GlossaryTooltipsApplicationCustomizer
  extends BaseApplicationCustomizer<IGlossaryTooltipsApplicationCustomizerProperties> {

  private popover: HTMLElement | null = null;
  private hideTimer: number | null = null;
  private termRegex: RegExp | null = null;
  private defMap: Map<string, string> = new Map();
  private contentObserver: MutationObserver | null = null;
  private wrapDebounce: number | null = null;
  private editModeWatcher: number | null = null;
  private isActivated = false;

  public onInit(): Promise<void> {
    const targetPath = (this.properties.pagePath || HANDBOOK_PAGE_PATH).toLowerCase();
    const currentPath = window.location.pathname.toLowerCase();
    const webUrl = this.context.pageContext.web.serverRelativeUrl;

    if (matchHandbookPath(currentPath, targetPath, webUrl) === undefined) {
      Log.info(LOG_SOURCE, `Skipping: ${currentPath} matches neither ${targetPath} nor the site root ${webUrl}`);
      return Promise.resolve();
    }

    // Bail out in edit mode (?Mode=Edit). This customizer walks the page
    // with a TreeWalker and wraps glossary terms in <span class="term">.
    // That DOM mutation interferes with the SP page editor — the editor
    // can't reliably edit text that we've just wrapped, and our
    // MutationObserver would re-wrap on every keystroke. Same approach
    // as the sibling HandbookExperience customizer.
    const mode = (new URLSearchParams(window.location.search).get('Mode') || '').toLowerCase();
    if (mode === 'edit') {
      Log.info(LOG_SOURCE, 'Edit mode (?Mode=Edit) detected — glossary tooltips disabled to leave the SP editor untouched.');
      return Promise.resolve();
    }

    Log.info(LOG_SOURCE, `Activating glossary tooltips on ${currentPath}`);

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => this.activate());
    } else {
      this.activate();
    }

    return Promise.resolve();
  }

  private activate(): void {
    // Cross-instance singleton guard. When debugging a customizer that is also
    // deployed, SharePoint runs both copies at once — two instances would each
    // append a #hb-term-popover and double-process the page. First one wins;
    // any later instance bails. (Resets on the edit→view reload.)
    const w = window as unknown as { __hbGlossaryActive?: boolean };
    if (w.__hbGlossaryActive) {
      Log.info(LOG_SOURCE, 'Another GlossaryTooltips instance is already active — skipping.');
      return;
    }
    w.__hbGlossaryActive = true;

    this.isActivated = true;
    this.startEditModeWatcher();
    this.injectFallbackStyles();

    this.popover = document.createElement('div');
    this.popover.id = POPOVER_ID;
    this.popover.className = 'hb-term-popover';
    this.popover.setAttribute('role', 'tooltip');
    document.body.appendChild(this.popover);

    document.addEventListener('mouseenter', this.onMouseEnter, true);
    document.addEventListener('mouseleave', this.onMouseLeave, true);
    document.addEventListener('focusin', this.onFocusIn);
    document.addEventListener('focusout', this.onFocusOut);
    document.addEventListener('click', this.onClick);
    window.addEventListener('scroll', this.hide, { passive: true });
    window.addEventListener('resize', this.hide);

    // Load list + start wrapping. Failures are silent — prototype-inline
    // .term spans still work without the list.
    console.log('[hb-glossary] activate() ran — fetching list...');
    mark('glossary', 'activate() entered');
    scheduleDump(12000);

    this.loadGlossary()
      .then(terms => {
        console.log(`[hb-glossary] Loaded ${terms.length} terms from list "${this.listTitle()}"`);
        mark('glossary', 'glossary terms ready', `${terms.length} terms`);
        if (terms.length > 0) {
          this.buildMatcher(terms);
          this.waitForContentThenWrap();
        }
      })
      .catch(err => {
        console.error('[hb-glossary] Glossary list load failed:', err);
        mark('glossary', '!!! glossary load FAILED', String(err));
      });
  }

  // Match the sibling HandbookExperience customizer: poll for ?Mode=Edit
  // after activation. SharePoint's modern editor often transitions
  // view → edit without a full reload, so the onInit-only check above
  // doesn't catch the live transition.
  //
  // view → edit: tear down inline (disconnect observer, hide popover,
  //              remove document-level listeners). Existing .term spans
  //              already in the DOM stay there — they're inert without
  //              the popover + listeners.
  // edit → view: reload, so the customizer re-runs against a clean DOM.
  private wasEditMode = false;
  private startEditModeWatcher(): void {
    if (this.editModeWatcher) window.clearInterval(this.editModeWatcher);
    this.editModeWatcher = window.setInterval(() => {
      const mode = (new URLSearchParams(window.location.search).get('Mode') || '').toLowerCase();
      const isEditMode = mode === 'edit';

      if (isEditMode && this.isActivated && !this.wasEditMode) {
        console.log('[hb-glossary] Mode=Edit detected — tearing down tooltip listeners.');
        this.tearDownForEditMode();
        this.wasEditMode = true;
      } else if (!isEditMode && this.wasEditMode) {
        console.log('[hb-glossary] Mode=Edit cleared — reloading to restore tooltips.');
        window.clearInterval(this.editModeWatcher!);
        this.editModeWatcher = null;
        window.location.reload();
      }
    }, 500);
  }

  private tearDownForEditMode(): void {
    if (this.contentObserver) {
      this.contentObserver.disconnect();
      this.contentObserver = null;
    }
    if (this.wrapDebounce) {
      window.clearTimeout(this.wrapDebounce);
      this.wrapDebounce = null;
    }
    if (this.popover) {
      this.popover.style.display = 'none';
    }
    document.removeEventListener('mouseenter', this.onMouseEnter, true);
    document.removeEventListener('mouseleave', this.onMouseLeave, true);
    document.removeEventListener('focusin', this.onFocusIn);
    document.removeEventListener('focusout', this.onFocusOut);
    document.removeEventListener('click', this.onClick);
    window.removeEventListener('scroll', this.hide);
    window.removeEventListener('resize', this.hide);
  }

  private listTitle(): string {
    return this.properties.listTitle || LIST_TITLE;
  }

  // ===== Glossary list loading =====
  private async loadGlossary(): Promise<IGlossaryTerm[]> {
    // Try cache first
    try {
      const raw = localStorage.getItem(CACHE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as { expires: number; terms: IGlossaryTerm[] };
        if (parsed.expires > Date.now() && Array.isArray(parsed.terms)) {
          mark('glossary', 'terms from CACHE (no network)');
          return parsed.terms;
        }
      }
    } catch { /* ignore */ }

    // Fetch from SharePoint REST
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const endpoint =
      `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(this.listTitle())}')` +
      `/items?$select=Title,Definition&$top=${MAX_TERMS}`;

    mark('glossary', 'REST fetch started (cache miss)');
    const fetchStart = performance.now();
    const res = await fetch(endpoint, {
      credentials: 'include',
      headers: { accept: 'application/json;odata=nometadata' }
    });
    mark('glossary', 'REST fetch returned', `took ${Math.round(performance.now() - fetchStart)}ms`);
    if (!res.ok) {
      throw new Error(`HTTP ${res.status} fetching glossary list`);
    }
    const data = await res.json() as { value: Array<{ Title?: string; Definition?: string }> };
    const terms: IGlossaryTerm[] = (data.value || [])
      .filter(it => !!it.Title && !!it.Definition)
      .map(it => ({
        term: String(it.Title).trim(),
        definition: String(it.Definition).trim()
      }));

    try {
      localStorage.setItem(CACHE_KEY, JSON.stringify({
        expires: Date.now() + CACHE_TTL_MS,
        terms
      }));
    } catch { /* quota etc — ignore */ }

    return terms;
  }

  // ===== Term-matching =====
  private buildMatcher(terms: IGlossaryTerm[]): void {
    // Longest term first so "Tier 1 instruction" wins over "Tier 1"
    const sorted = [...terms].sort((a, b) => b.term.length - a.term.length);
    const escaped = sorted.map(t => t.term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    // Input strings are escaped above, so the constructed regex is safe.
    // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
    this.termRegex = new RegExp(`\\b(${escaped.join('|')})\\b`, 'i');

    this.defMap.clear();
    for (const t of sorted) {
      this.defMap.set(t.term.toLowerCase(), t.definition);
    }
  }

  // Wait for the SharePoint content area to render, then wrap and set up
  // a MutationObserver that re-wraps when editors change content.
  private waitForContentThenWrap(retries = 20): void {
    const content = document.querySelector<HTMLElement>(
      '[data-automation-id="mainScrollRegionInnerContent"], [data-automation-id="contentScrollRegion"]'
    );
    if (content) {
      // Time the first wrap. This is a TreeWalker over every text node in a
      // ~290KB document, so it is the glossary's single biggest cost and a prime
      // suspect for main-thread contention right after the chrome appears.
      const t0 = performance.now();
      this.wrapAllTerms(content);
      mark('glossary', 'first wrap pass complete', `took ${Math.round(performance.now() - t0)}ms`);

      this.attachContentObserver(content);
      window.setTimeout(() => this.wrapAllTerms(content), 800);
      return;
    }
    if (retries > 0) {
      window.setTimeout(() => this.waitForContentThenWrap(retries - 1), 250);
    } else {
      mark('glossary', '!!! GAVE UP waiting for SP canvas', '20 retries exhausted — no terms wrapped');
    }
  }

  // Re-wrap churn control.
  //
  // Measured on the live page (v1.0.8.0 instrumentation): the observer fired 8
  // full re-wraps, still going at 35 SECONDS after load, 20–90ms each. Each one
  // is a TreeWalker over every text node in the document, running while someone
  // is trying to read.
  //
  // The cause is NOT a feedback loop from our own mutations — wrapAllTerms
  // already disconnects before it edits and reattaches after. It is SharePoint
  // itself: this page has 123 canvas controls and SP keeps lazily rendering and
  // re-rendering them long after load, and every one of those mutations used to
  // buy a full document re-walk.
  //
  // Two changes:
  //
  //   1. Stop watching characterData. New glossary-eligible content always
  //      arrives as new ELEMENTS (childList) — a text node whose value changes
  //      in place can't introduce a term we haven't already considered. SP
  //      re-renders text nodes constantly, so this was the noisiest trigger and
  //      the least useful.
  //   2. Settle and stop. wrapAllTerms now reports how many NEW occurrences it
  //      wrapped. Three consecutive passes that wrap nothing means the page has
  //      stopped producing new content, so we disconnect for good. Any pass that
  //      does find something resets the counter, so a genuinely late-rendering
  //      section is still caught.
  //
  // Disconnecting permanently is safe here: the only way content changes after
  // that point is a page edit, and the edit-mode watcher reloads the page on the
  // way back out of edit mode, which re-runs everything from scratch.
  private rewrapCount = 0;
  private zeroYieldPasses = 0;
  private static readonly SETTLE_AFTER_ZERO_YIELD_PASSES = 3;

  private attachContentObserver(target: HTMLElement): void {
    if (this.contentObserver) this.contentObserver.disconnect();
    this.contentObserver = new MutationObserver(() => {
      if (this.wrapDebounce) window.clearTimeout(this.wrapDebounce);
      this.wrapDebounce = window.setTimeout(() => {
        this.rewrapCount += 1;
        const t0 = performance.now();

        // NB: wrapAllTerms reattaches the observer on its way out, so the
        // settle check below has to run AFTER it, or it would be undone.
        const wrapped = this.wrapAllTerms(target);
        const tookMs = Math.round(performance.now() - t0);

        if (wrapped > 0) {
          this.zeroYieldPasses = 0;
        } else {
          this.zeroYieldPasses += 1;
        }

        if (this.rewrapCount <= 10) {
          mark(
            'glossary',
            `re-wrap #${this.rewrapCount} (observer)`,
            `took ${tookMs}ms, wrapped ${wrapped}, zero-yield streak ${this.zeroYieldPasses}`
          );
        }

        if (this.zeroYieldPasses >= GlossaryTooltipsApplicationCustomizer.SETTLE_AFTER_ZERO_YIELD_PASSES) {
          if (this.contentObserver) {
            this.contentObserver.disconnect();
            this.contentObserver = null;
          }
          if (this.wrapDebounce) {
            window.clearTimeout(this.wrapDebounce);
            this.wrapDebounce = null;
          }
          console.log(`[hb-glossary] Settled after ${this.rewrapCount} re-wraps — observer disconnected.`);
          mark('glossary', 'settled — observer disconnected', `after ${this.rewrapCount} re-wraps`);
        }
      }, 800);
    });
    this.contentObserver.observe(target, {
      childList: true,
      subtree: true
    });
  }

  // Walk text nodes inside the content area and wrap matching glossary terms
  // in <span class="term" data-def="…">. Terms are de-duplicated PER BLOCK:
  // a given term is highlighted only on its FIRST occurrence within the same
  // heading-bounded section. Every H1–H6 starts a new block, so each FAQ
  // question (an H4 and the body that follows it, until the next heading) is
  // its own block. If "assessment" appears four times under one FAQ answer,
  // only the first gets a tooltip.
  //
  // The per-block "already highlighted" set is seeded from .term spans that
  // are already in the DOM, so the periodic re-wrap passes and the
  // MutationObserver don't re-introduce duplicates on later occurrences.
  //
  // NOTE: this de-dup operates entirely within this (Glossary) customizer's
  // content scope and never touches the sidebar/reader DOM — the menu is
  // built by the separate HandbookExperience customizer from headings.
  //
  // We pause the MutationObserver during DOM mutations to avoid feedback loops.
  // Returns the number of NEW occurrences wrapped, so attachContentObserver can
  // tell a productive pass from a wasted one and settle once the page goes quiet.
  private wrapAllTerms(scope: HTMLElement): number {
    if (!this.termRegex) return 0;

    const wasObserving = !!this.contentObserver;
    if (this.contentObserver) this.contentObserver.disconnect();

    // Skip text inside links, code blocks, our own FAQ summaries, editor UI,
    // and SP list web parts (otherwise term definitions inside the glossary
    // list itself would get self-referentially wrapped).
    const SKIP_SELECTOR = [
      'a', 'code', 'pre', 'script', 'style',
      '.term', '.hb-faq-summary',
      '[data-hb]',
      '[contenteditable="true"]',
      '[data-sp-feature-tag*="List" i]',
      '[data-automation-id*="list" i]',
      '[role="grid"]',
      '[role="row"]',
      '[role="rowgroup"]',
      '.ms-DetailsList',
      '[class*="DetailsRow"]',
      '[class*="listCardContainer"]',
      // SharePoint "Call to Action" web part cards — don't auto-wrap terms
      // inside the short title/description of these tile links
      '[class*="overlayContainer-"]',
      '[class*="htmlOverlayTextWrapper-"]',
      '[data-automation-id="image-overlay-text"]'
    ].join(', ');

    // Per-block set of term keys (lowercased) already highlighted. The block
    // is keyed by the heading element that opens it; content before the first
    // heading is keyed by the scope element itself.
    const seenByBlock = new Map<Node, Set<string>>();
    const seenSetFor = (block: Node): Set<string> => {
      let set = seenByBlock.get(block);
      if (!set) { set = new Set<string>(); seenByBlock.set(block, set); }
      return set;
    };

    // Single document-order walk that visits headings (block boundaries),
    // existing .term spans (to seed the seen-set), and candidate text nodes.
    const walker = document.createTreeWalker(
      scope,
      NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT,
      {
        acceptNode: (node: Node): number => {
          if (node.nodeType === Node.ELEMENT_NODE) {
            const el = node as Element;
            // Prune skip-subtrees entirely — but never prune .term, whose
            // text we still want to read to seed the per-block seen-set.
            if (!el.classList.contains('term') && el.matches(SKIP_SELECTOR)) {
              return NodeFilter.FILTER_REJECT;
            }
            // Surface headings and existing terms; descend through the rest.
            if (/^H[1-6]$/.test(el.tagName) || el.classList.contains('term')) {
              return NodeFilter.FILTER_ACCEPT;
            }
            return NodeFilter.FILTER_SKIP;
          }
          const parent = (node as Text).parentElement;
          if (!parent) return NodeFilter.FILTER_REJECT;
          if (parent.closest(SKIP_SELECTOR)) return NodeFilter.FILTER_REJECT;
          const text = node.nodeValue || '';
          if (text.length < 3) return NodeFilter.FILTER_REJECT;
          if (!this.termRegex!.test(text)) return NodeFilter.FILTER_REJECT;
          return NodeFilter.FILTER_ACCEPT;
        }
      } as NodeFilter
    );

    // Resolve the de-dup block for a node. De-dup is PER FAQ: once the handbook
    // customizer wires collapsibles, each FAQ's H4 carries .hb-faq-summary,
    // which this customizer's SKIP_SELECTOR prunes — so the heading walk below
    // never sees the H4 and would otherwise lump every FAQ in a section into a
    // single block (a term highlighted in the first FAQ then suppressed in all
    // the rest). Anchoring on the enclosing .hb-faq wrapper (which is NOT
    // skipped) restores true once-per-FAQ highlighting. Content outside any FAQ
    // falls back to the nearest heading block.
    const blockFor = (el: Element | null, headingBlock: Node): Node =>
      (el && el.closest('.hb-faq')) || headingBlock;

    const candidates: Array<{ node: Text; block: Node }> = [];
    let currentBlock: Node = scope;
    let n: Node | null;
    while ((n = walker.nextNode())) {
      if (n.nodeType === Node.ELEMENT_NODE) {
        const el = n as Element;
        if (/^H[1-6]$/.test(el.tagName)) {
          currentBlock = el;                       // open a new heading block
        } else if (el.classList.contains('term')) {
          // Seed from a prior wrap, keyed by the same block we'd wrap into.
          const key = (el.textContent || '').trim().toLowerCase();
          if (key) seenSetFor(blockFor(el, currentBlock)).add(key);
        }
      } else {
        candidates.push({
          node: n as Text,
          block: blockFor((n as Text).parentElement, currentBlock)
        });
      }
    }

    let wrappedCount = 0;
    for (const { node, block } of candidates) {
      if (this.wrapFirstUnseenMatch(node, seenSetFor(block))) wrappedCount++;
    }
    console.log(`[hb-glossary] Wrapped ${wrappedCount} new occurrences (${candidates.length} candidate text nodes, ${seenByBlock.size} blocks)`);

    if (wasObserving) {
      this.attachContentObserver(scope);
    }

    return wrappedCount;
  }

  // Wrap the first regex match in this text node whose term has NOT already
  // been highlighted in the given block, then record it as seen. Returns true
  // if a wrap happened. Skipping already-seen leading matches lets a later,
  // not-yet-seen term in the same node still get its tooltip.
  private wrapFirstUnseenMatch(node: Text, seenInBlock: Set<string>): boolean {
    if (!this.termRegex) return false;
    const text = node.nodeValue || '';

    // Global copy so we can scan every match in the node (the shared
    // this.termRegex is non-global and used statelessly for .test()).
    // Source is built from escaped list values, so the pattern is safe.
    // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
    const re = new RegExp(this.termRegex.source, 'gi');
    let match: RegExpExecArray | null;
    while ((match = re.exec(text))) {
      const matchText = match[0];
      const key = matchText.toLowerCase();
      if (seenInBlock.has(key)) continue;          // already highlighted here
      const definition = this.defMap.get(key);
      if (!definition) continue;

      const parent = node.parentNode;
      if (!parent) return false;

      const before = text.substring(0, match.index);
      const after = text.substring(match.index + matchText.length);

      if (before) parent.insertBefore(document.createTextNode(before), node);

      const span = document.createElement('span');
      span.className = 'term';
      span.setAttribute('tabindex', '0');
      span.setAttribute('data-def', definition);
      span.textContent = matchText;
      parent.insertBefore(span, node);

      if (after) parent.insertBefore(document.createTextNode(after), node);

      parent.removeChild(node);
      seenInBlock.add(key);
      return true;
    }
    return false;
  }

  // The popover styles live in HandbookExperience.module.scss. Inject a minimal
  // fallback so the tooltip still works if that customizer is not deployed.
  private injectFallbackStyles(): void {
    if (document.getElementById('hb-term-popover-styles')) return;
    const style = document.createElement('style');
    style.id = 'hb-term-popover-styles';
    style.textContent = `
      .hb-term-popover {
        display: none; position: fixed; z-index: 9999;
        background: #002a42; color: #fff; padding: 0.75rem 1rem;
        border-radius: 8px; font-size: 0.82rem; line-height: 1.55;
        width: max-content; max-width: 360px;
        box-shadow: 0 6px 24px rgba(0,0,0,0.25); pointer-events: none;
      }
      .hb-term-popover.visible { display: block; }
    `;
    document.head.appendChild(style);
  }

  private isTerm(el: EventTarget | null): el is HTMLElement {
    return !!(el && (el as HTMLElement).classList?.contains('term'));
  }

  private onMouseEnter = (e: Event): void => {
    if (this.isTerm(e.target)) {
      if (this.hideTimer) { window.clearTimeout(this.hideTimer); this.hideTimer = null; }
      this.show(e.target);
    }
  };

  private onMouseLeave = (e: Event): void => {
    if (this.isTerm(e.target)) {
      this.hideTimer = window.setTimeout(() => this.hide(), 120);
    }
  };

  private onFocusIn = (e: FocusEvent): void => {
    if (this.isTerm(e.target)) this.show(e.target);
  };

  private onFocusOut = (e: FocusEvent): void => {
    if (this.isTerm(e.target)) this.hide();
  };

  private onClick = (e: MouseEvent): void => {
    if (this.isTerm(e.target)) {
      if (this.popover?.classList.contains('visible')) {
        this.hide();
      } else {
        this.show(e.target);
      }
    } else if (this.popover && !this.popover.contains(e.target as Node)) {
      this.hide();
    }
  };

  private show(term: HTMLElement): void {
    if (!this.popover) return;
    const def = term.getAttribute('data-def');
    if (!def) return;

    this.popover.textContent = def;
    this.popover.className = 'hb-term-popover visible';

    const rect = term.getBoundingClientRect();
    const pw = this.popover.offsetWidth;
    const ph = this.popover.offsetHeight;

    let left = rect.left + rect.width / 2 - pw / 2;
    left = Math.max(8, Math.min(left, window.innerWidth - pw - 8));

    let top: number;
    if (rect.top - ph - 12 > 0) {
      top = rect.top - ph - 10;
      this.popover.classList.add('arrow-bottom');
    } else {
      top = rect.bottom + 10;
      this.popover.classList.add('arrow-top');
    }

    this.popover.style.left = left + 'px';
    this.popover.style.top = top + 'px';
  }

  private hide = (): void => {
    if (!this.popover) return;
    this.popover.className = 'hb-term-popover';
    this.popover.style.left = '-9999px';
  };
}
