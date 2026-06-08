import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import './HandbookExperience.module.scss';

const LOG_SOURCE = 'HandbookExperienceApplicationCustomizer';
const HANDBOOK_PAGE_PATH = '/sitepages/home1.aspx';

const FONTS: Record<string, string> = {
  sans: "'Libre Franklin','Franklin Gothic','ITC Franklin Gothic','Helvetica Neue',Helvetica,Arial,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",
  serif: "Georgia,'Times New Roman',serif",
  dyslexic: "'OpenDyslexic','Comic Sans MS',sans-serif"
};

const buildToolbarHtml = (logoUrl: string): string => `
  <header class="toolbar" role="banner">
    <div class="toolbar-left">
      <img src="${logoUrl}" alt="HKIS Logo" class="toolbar-logo" onerror="this.style.display='none'">
      <span class="toolbar-brand">HKIS Teaching &amp; Learning Handbook</span>
    </div>
    <div class="toolbar-right">
      <div class="search-wrap">
        <button class="tb" id="btn-search" data-hb="search-toggle" aria-label="Search" title="Search handbook">&#128269;</button>
        <input type="text" class="search-input" id="search-input" placeholder="Search handbook..." aria-label="Search handbook" autocomplete="off" />
        <div class="search-results" id="search-results"></div>
      </div>
      <button class="tb" id="btn-agents" data-hb="agents-toggle" aria-label="Ask a Question via SharePoint Agents" title="Ask a Question">&#10024; Ask a Question</button>
      <div class="aa-wrap">
        <button class="tb" id="btn-aa" data-hb="aa-toggle" aria-label="Reading settings" title="Reading settings">Aa</button>
        <div class="aa-popover" id="aa-popover">
          <div class="aa-section">
            <div class="aa-label">Text size</div>
            <div class="aa-size-row">
              <button class="tb" data-hb="font-smaller" aria-label="Smaller text">A&minus;</button>
              <span class="aa-size-label" id="size-label">100%</span>
              <button class="tb" data-hb="font-larger" aria-label="Larger text">A+</button>
            </div>
            <div style="margin-top:0.4rem;text-align:center;">
              <button class="tb" data-hb="font-reset" style="font-size:0.72rem;">Reset</button>
            </div>
          </div>
          <div class="aa-section">
            <div class="aa-label">Font</div>
            <div class="aa-row">
              <button class="tb" id="btn-font-sans" data-hb="font" data-font="sans">Sans</button>
              <button class="tb" id="btn-font-serif" data-hb="font" data-font="serif">Serif</button>
              <button class="tb" id="btn-font-dyslexic" data-hb="font" data-font="dyslexic">Dyslexic</button>
            </div>
          </div>
          <div class="aa-section">
            <div class="aa-label">Theme</div>
            <div class="aa-themes">
              <div class="aa-theme-opt">
                <button class="theme-swatch swatch-light" id="sw-light" data-hb="theme" data-theme="light" aria-label="Light theme"></button>
                <span>Light</span>
              </div>
              <div class="aa-theme-opt">
                <button class="theme-swatch swatch-sepia" id="sw-sepia" data-hb="theme" data-theme="sepia" aria-label="Sepia theme"></button>
                <span>Sepia</span>
              </div>
              <div class="aa-theme-opt">
                <button class="theme-swatch swatch-dark" id="sw-dark" data-hb="theme" data-theme="dark" aria-label="Dark theme"></button>
                <span>Dark</span>
              </div>
            </div>
          </div>
        </div>
      </div>
      <button class="tb" id="btn-menu" data-hb="menu-toggle" aria-label="Toggle navigation" title="Sections">&#9776;</button>
    </div>
  </header>
  <div class="progress-bar"><div class="progress-bar-fill" id="progress-fill"></div></div>
`;

interface ISearchHit {
  title: string;
  text: string;
  el: HTMLElement;
  type: 'section' | 'question' | 'glossary';
  score: number;
}

export interface IHandbookExperienceApplicationCustomizerProperties {
  pagePath?: string;
}

export default class HandbookExperienceApplicationCustomizer
  extends BaseApplicationCustomizer<IHandbookExperienceApplicationCustomizerProperties> {

  private toolbarRendered = false;
  private currentScale = 1;
  private searchIndex: ISearchHit[] | null = null;
  private prevHighlights: HTMLElement[] = [];
  private activated = false;
  private contentObserver: MutationObserver | null = null;
  private editModeWatcher: number | null = null;
  private scrollSpyObserver: IntersectionObserver | null = null;
  private rebuildTimer: number | null = null;
  private scrollUpdateFn: (() => void) | null = null;
  private scrollContainer: HTMLElement | null = null;
  private initialHashConsumed = false;

  public onInit(): Promise<void> {
    const targetPath = (this.properties.pagePath || HANDBOOK_PAGE_PATH).toLowerCase();
    const currentPath = window.location.pathname.toLowerCase();

    if (currentPath.indexOf(targetPath, currentPath.length - targetPath.length) === -1) {
      Log.info(LOG_SOURCE, `Skipping: ${currentPath} does not match ${targetPath}`);
      return Promise.resolve();
    }

    // Bail out entirely in edit mode.
    //
    // Our content-shaping operations (wireCollapsibles re-wrapping H4 siblings
    // into .hb-faq-body, wireCtaCards intercepting clicks on overlay
    // containers, the MutationObserver-driven rebuild on every DOM change)
    // fight the SharePoint page editor for control of the same DOM tree.
    // Symptoms:
    //   - Scroll jumps back to a previous heading while editing further down
    //     the page (editor restores its last anchor when our rewrite yanks
    //     the content tree out from under it)
    //   - SP command-bar icons (New / Page details / Preview / Edit) render
    //     as missing-glyph boxes (our SCSS rules under .hb-active reach into
    //     edit-mode chrome that didn't exist in view mode)
    //
    // SharePoint signals edit mode via `?Mode=Edit` in the URL. View-mode
    // navigation lands on `?Mode=View` or no `Mode` param at all. When the
    // operator clicks "Edit" in the SP page action bar, SP does a full page
    // navigation to ?Mode=Edit, so this check is evaluated at the right
    // moment — onInit fires fresh and we no-op cleanly.
    const mode = (new URLSearchParams(window.location.search).get('Mode') || '').toLowerCase();
    if (mode === 'edit') {
      Log.info(LOG_SOURCE, 'Edit mode (?Mode=Edit) detected — handbook experience disabled to leave the SP editor untouched.');
      return Promise.resolve();
    }

    Log.info(LOG_SOURCE, `Activating handbook experience on ${currentPath}`);

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => this.activate());
    } else {
      this.activate();
    }

    return Promise.resolve();
  }

  // ===== Preferences =====
  private getPref(k: string, def: string): string {
    try { return localStorage.getItem('hb_' + k) || def; } catch { return def; }
  }
  private setPref(k: string, v: string): void {
    try { localStorage.setItem('hb_' + k, v); } catch { /* no-op */ }
  }

  // Load Libre Franklin from Google Fonts. Idempotent.
  private loadFonts(): void {
    if (!document.getElementById('hb-font-pc1')) {
      const pc1 = document.createElement('link');
      pc1.id = 'hb-font-pc1';
      pc1.rel = 'preconnect';
      pc1.href = 'https://fonts.googleapis.com';
      document.head.appendChild(pc1);
    }
    if (!document.getElementById('hb-font-pc2')) {
      const pc2 = document.createElement('link');
      pc2.id = 'hb-font-pc2';
      pc2.rel = 'preconnect';
      pc2.href = 'https://fonts.gstatic.com';
      pc2.setAttribute('crossorigin', '');
      document.head.appendChild(pc2);
    }
    if (!document.getElementById('hb-font-libre-franklin')) {
      const link = document.createElement('link');
      link.id = 'hb-font-libre-franklin';
      link.rel = 'stylesheet';
      link.href = 'https://fonts.googleapis.com/css2?family=Libre+Franklin:wght@300;400;500;600;700;800&display=swap';
      document.head.appendChild(link);
    }
  }

  // ===== Activation =====
  private activate(): void {
    if (this.activated) return;

    // Cross-instance singleton guard. SharePoint runs BOTH the deployed copy of
    // this customizer AND a debug copy when the page is loaded with
    // ?debugManifestsFile (i.e. when you debug a customizer that is also
    // deployed). Two live instances each inject their own toolbar + sidebar +
    // reading-pos, producing duplicate chrome that fights — e.g. two #sidebar
    // nodes, one populated (35 links) and one empty, with the empty one on top.
    // JS is single-threaded and activate() is synchronous, so the first
    // instance to reach here flips a window-level flag and any later instance
    // bails before injecting anything. (The flag resets on the edit→view reload.)
    const w = window as unknown as { __hbExperienceActive?: boolean };
    if (w.__hbExperienceActive) {
      Log.info(LOG_SOURCE, 'Another HandbookExperience instance is already active — skipping to avoid duplicate chrome.');
      return;
    }
    w.__hbExperienceActive = true;

    this.activated = true;

    this.loadFonts();

    const root = document.documentElement;
    root.classList.add('hb-active');

    const preferredDark =
      typeof window.matchMedia === 'function' &&
      window.matchMedia('(prefers-color-scheme: dark)').matches;
    const theme = this.getPref('theme', preferredDark ? 'dark' : 'light');
    root.setAttribute('data-theme', theme);

    this.currentScale = parseFloat(this.getPref('fontScale', '1')) || 1;
    root.style.setProperty('--font-scale', String(this.currentScale));

    this.applyFont(this.getPref('font', 'sans'), false);

    this.injectToolbar();
    this.injectSidebar();
    this.injectOverlay();
    this.injectReadingPos();
    this.injectAgentsFab();

    // When embedded (Teams tab, iframe), Copilot isn't reachable in-context.
    // Update both Agents button tooltips so users know what the click does.
    if (this.isEmbeddedContext()) {
      const tbBtn = document.getElementById('btn-agents');
      const fab = document.getElementById('hb-agents-fab');
      const hint = 'Opens the full handbook in a new tab so you can ask Copilot';
      tbBtn?.setAttribute('title', hint);
      fab?.setAttribute('title', hint);
    }

    this.wireGlobalClicks();
    this.wireSidebar();
    this.wireKeyboardShortcuts();

    // Build TOC + scroll spy. SP may still be hydrating the canvas when we
    // first run, so we both (a) attempt immediately, (b) retry on a few
    // delays, and (c) attach a MutationObserver so any later DOM changes
    // (e.g., editors adding/removing sections) trigger an automatic rebuild.
    this.waitForContentThenWatch();

    this.startEditModeWatcher();
  }

  // ===== Edit-mode watcher =====
  // SharePoint's modern editor transitions a page from view → edit by
  // mutating the URL to ?Mode=Edit without a full page reload. The
  // onInit-only check at the top of this file therefore doesn't fire
  // during a live transition — the customizer was activated for view
  // mode and stays active as the editor mounts.
  //
  // Symptoms when that happens:
  //   - Scroll jumps to a previous heading while editing further down
  //     the page (wireCollapsibles re-wraps siblings of H4s every time
  //     the editor mutates the DOM, yanking the editor's caret position)
  //   - SP command-bar icons render as missing-glyph boxes (CSS rules
  //     scoped under .hb-active reach into edit-mode chrome)
  //   - Sidebar TOC rebuilds constantly off headings the editor is
  //     restructuring, which surfaces as visible flicker
  //
  // Fix: poll the URL every 500ms after activation. Cross either
  // direction (view→edit or edit→view) and react:
  //   - view → edit: hide the handbook chrome inline (no reload flash),
  //                  disconnect every observer + scroll listener so we
  //                  stop fighting the editor for DOM control
  //   - edit → view: reload the page so the customizer re-activates
  //                  against a freshly-rendered (editor-cleaned-up) DOM.
  //                  The reload here is fine UX-wise — the user just
  //                  clicked Publish or X'd out of edit mode, they
  //                  expect a fresh page anyway.
  //
  // Polling instead of pushState/popstate hooks because SP's internal
  // SPA navigation doesn't reliably emit those events; 500ms is cheap
  // and catches every transition within half a second.
  private wasEditMode = false;
  private startEditModeWatcher(): void {
    if (this.editModeWatcher) window.clearInterval(this.editModeWatcher);
    this.editModeWatcher = window.setInterval(() => {
      const mode = (new URLSearchParams(window.location.search).get('Mode') || '').toLowerCase();
      const isEditMode = mode === 'edit';

      if (isEditMode && this.activated && !this.wasEditMode) {
        Log.info(LOG_SOURCE, 'Mode=Edit detected — tearing down handbook chrome (no reload).');
        this.tearDownForEditMode();
        this.wasEditMode = true;
      } else if (!isEditMode && this.wasEditMode) {
        Log.info(LOG_SOURCE, 'Mode=Edit cleared — reloading to restore handbook chrome.');
        window.clearInterval(this.editModeWatcher!);
        this.editModeWatcher = null;
        window.location.reload();
      }
    }, 500);
  }

  // Inline tear-down — invoked when the URL transitions to ?Mode=Edit
  // post-activation. Strips the `.hb-active` class so all our CSS goes
  // dormant, hides every injected DOM node so SP's editor canvas is
  // unobstructed, and disconnects every observer / listener so the
  // editor can mutate the DOM without us re-wrapping headings under it.
  //
  // We do NOT unwrap the .hb-faq collapsibles or the .term glossary
  // spans inserted earlier — those are inert DOM once .hb-active is
  // gone, and unwrapping them would require traversing the page during
  // the very moment the editor is mounting. Leaving them in is empirically
  // OK; SP's editor edits the content inside them transparently.
  private tearDownForEditMode(): void {
    document.documentElement.classList.remove('hb-active');

    const idsToHide = ['hb-toolbar-root', 'sidebar', 'overlay', 'hb-agents-fab', 'reading-pos'];
    idsToHide.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = 'none';
    });

    if (this.contentObserver) {
      this.contentObserver.disconnect();
      this.contentObserver = null;
    }
    if (this.scrollSpyObserver) {
      this.scrollSpyObserver.disconnect();
      this.scrollSpyObserver = null;
    }
    if (this.scrollUpdateFn) {
      const target: EventTarget = this.scrollContainer || window;
      target.removeEventListener('scroll', this.scrollUpdateFn);
      this.scrollUpdateFn = null;
    }
    if (this.rebuildTimer) {
      window.clearTimeout(this.rebuildTimer);
      this.rebuildTimer = null;
    }
  }

  // ===== Toolbar injected directly into body at high z-index =====
  private injectToolbar(): void {
    if (this.toolbarRendered) return;
    const root = document.createElement('div');
    root.id = 'hb-toolbar-root';
    const logoUrl = `${this.context.pageContext.web.serverRelativeUrl}/SiteAssets/HKIS-logo.png`;
    root.innerHTML = buildToolbarHtml(logoUrl);
    document.body.insertBefore(root, document.body.firstChild);
    this.toolbarRendered = true;

    // Direct click handler on the toolbar root — more reliable than body
    // delegation, which SharePoint occasionally swallows.
    root.addEventListener('click', (e) => {
      const target = e.target as HTMLElement | null;
      const hb = target?.closest<HTMLElement>('[data-hb]');
      if (!hb) return;
      e.stopPropagation();
      Log.info(LOG_SOURCE, `Toolbar action: ${hb.dataset.hb}`);
      this.handleAction(hb.dataset.hb || '', hb);
    });

    this.markActiveTheme(this.getPref('theme', 'light'));
    this.markActiveFont(this.getPref('font', 'sans'));
    this.updateSizeLabel();
    this.wireSearchInput();
  }

  // ===== Injected non-toolbar DOM =====
  private injectSidebar(): void {
    const sidebar = document.createElement('nav');
    sidebar.className = 'sidebar';
    sidebar.id = 'sidebar';
    sidebar.setAttribute('role', 'navigation');
    sidebar.setAttribute('aria-label', 'Handbook navigation');
    // Default closed — user opens via the menu-toggle button in the toolbar.
    document.body.appendChild(sidebar);
  }

  private injectOverlay(): void {
    const overlay = document.createElement('div');
    overlay.className = 'overlay';
    overlay.id = 'overlay';
    overlay.setAttribute('data-hb', 'overlay');
    document.body.appendChild(overlay);
  }

  // Floating action button bottom-right that opens SP's Agents panel.
  // Same data-hb action as the toolbar's Agents button, so the existing
  // click delegation handles it.
  private injectAgentsFab(): void {
    const fab = document.createElement('button');
    fab.id = 'hb-agents-fab';
    fab.className = 'hb-agents-fab';
    fab.setAttribute('data-hb', 'agents-toggle');
    fab.setAttribute('aria-label', 'Ask a Question via SharePoint Agents');
    fab.setAttribute('title', 'Ask a Question');
    fab.innerHTML =
      '<span class="hb-agents-fab-icon">&#10024;</span>' +
      '<span class="hb-agents-fab-label">Ask a Question</span>';
    document.body.appendChild(fab);
  }

  private injectReadingPos(): void {
    const readingPos = document.createElement('div');
    readingPos.className = 'reading-pos';
    readingPos.id = 'reading-pos';
    readingPos.innerHTML = `
      <span><span class="reading-pos-section" id="rp-section"></span> &middot;
        Section <span id="rp-num">1</span> of <span id="rp-total">1</span></span>
      <div class="reading-pos-progress">
        <span id="rp-pct">0%</span>
        <div class="reading-pos-bar"><div class="reading-pos-bar-fill" id="rp-bar-fill"></div></div>
      </div>
    `;
    document.body.appendChild(readingPos);
  }

  // Wait for SP to render the canvas, then build TOC and attach an observer.
  private waitForContentThenWatch(retries = 20): void {
    const content = document.querySelector(
      '[data-automation-id="mainScrollRegionInnerContent"], [data-automation-id="contentScrollRegion"]'
    ) as HTMLElement | null;

    if (content) {
      this.rebuildSidebarAndSpy();
      this.attachContentObserver(content);
      // A second build a moment later catches any late-rendering collapsibles.
      window.setTimeout(() => this.rebuildSidebarAndSpy(), 800);
      return;
    }
    if (retries > 0) {
      window.setTimeout(() => this.waitForContentThenWatch(retries - 1), 250);
    }
  }

  private attachContentObserver(target: HTMLElement): void {
    if (this.contentObserver) this.contentObserver.disconnect();
    this.contentObserver = new MutationObserver(() => {
      if (this.rebuildTimer) window.clearTimeout(this.rebuildTimer);
      this.rebuildTimer = window.setTimeout(() => {
        this.rebuildSidebarAndSpy();
        // Invalidate cached search index so the next open rebuilds it.
        this.searchIndex = null;
      }, 400);
    });
    this.contentObserver.observe(target, {
      childList: true,
      subtree: true,
      characterData: false
    });
  }

  private rebuildSidebarAndSpy(): void {
    this.wireCollapsibles();
    this.wireCtaCards();
    this.buildSidebar();
    this.setupScrollSpy();
  }

  // Make SharePoint "Call to Action" cards fully clickable. The web part
  // renders text + a button — we listen for clicks anywhere on the card and
  // delegate to the inner button so the whole tile is the active hit target.
  private wireCtaCards(): void {
    document.querySelectorAll<HTMLElement>('[class*="overlayContainer-"]').forEach(card => {
      if (card.dataset.hbCtaWired === 'true') return;
      card.dataset.hbCtaWired = 'true';
      card.addEventListener('click', (e: MouseEvent) => {
        const target = e.target as HTMLElement | null;
        // If the click was on the button itself or any inner link/button,
        // let SP handle it natively. Otherwise, find and trigger the button.
        if (target?.closest('a, button, [class*="buttonContainer-"], [class*="actionContainer-"]')) {
          return;
        }
        const btn = card.querySelector<HTMLElement>(
          '[class*="buttonContainer-"] button, [class*="buttonContainer-"] a, ' +
          '[class*="actionContainer-"] button, [class*="actionContainer-"] a'
        );
        btn?.click();
      });
    });
  }

  // Transform every <h4> inside SP text content into a clickable FAQ
  // collapsible.
  //
  // SP's CKEditor wraps every heading in <div class="anchor-h4"> (it adds
  // permalink decoration spans next to the heading). Answer paragraphs are
  // SIBLINGS of that wrapper, not siblings of the H4 directly — so we walk
  // siblings of the H4's block container, not the H4.
  private wireCollapsibles(): void {
    const wasObserving = !!this.contentObserver;
    if (this.contentObserver) {
      this.contentObserver.disconnect();
    }

    const h4s = Array.from(document.querySelectorAll<HTMLHeadingElement>(
      '[data-automation-id="textBox"] h4, .ck-content h4, .rte-webpart h4'
    ));

    for (const h4 of h4s) {
      if (h4.closest('.hb-faq')) continue;

      // Find the H4's block-level container. SP uses .anchor-h4; fall back
      // to the H4 itself if SP didn't wrap it (e.g., older versions or
      // editor-flat content).
      const block: HTMLElement = (h4.closest('.anchor-h4') as HTMLElement) || h4;
      if (!block.parentNode) continue;

      // Collect siblings of the block until we hit another heading-containing
      // block (which marks the start of the next FAQ or section).
      const body: Element[] = [];
      let next = block.nextElementSibling;
      while (next) {
        if (/^H[1-6]$/.test(next.tagName)) break;
        if (next.querySelector(':scope > h1, :scope > h2, :scope > h3, :scope > h4, :scope > h5, :scope > h6')) break;
        const after = next.nextElementSibling;
        body.push(next);
        next = after;
      }

      // Absorb a media EMBED web part (YouTube / video — renders as an
      // <iframe> or <video>) that an author places directly beneath the H4's
      // text web part, so it collapses/expands WITH the FAQ. Deliberately
      // conservative: SharePoint interleaves the FAQ text web parts with other
      // content (empty spacer web parts, button/link web parts, the next
      // section's text). We ONLY pull in actual embeds, step over genuinely-
      // empty spacer web parts, and STOP at the first real content web part —
      // so we never swallow section content into a collapsed FAQ.
      const enclosingCc = h4.closest<HTMLElement>('[data-automation-id="CanvasControl"]');
      if (enclosingCc) {
        let nextCc = enclosingCc.nextElementSibling as HTMLElement | null;
        while (nextCc) {
          if (!nextCc.matches('[data-automation-id="CanvasControl"]')) {
            nextCc = nextCc.nextElementSibling as HTMLElement | null;
            continue;
          }
          if (nextCc.querySelector('h1, h2, h3, h4')) break;          // next FAQ/section
          if (nextCc.querySelector('iframe, video')) {                 // media embed → pull in
            const afterCc = nextCc.nextElementSibling as HTMLElement | null;
            body.push(nextCc);
            nextCc = afterCc;
            continue;
          }
          const isEmptySpacer =
            !nextCc.querySelector('img, [data-automation-id="imageWebPart"]')
            && (nextCc.textContent || '').trim() === '';
          if (isEmptySpacer) {                                         // empty spacer → step over
            nextCc = nextCc.nextElementSibling as HTMLElement | null;
            continue;
          }
          break;                                                       // real content → leave it
        }
      }

      // Wrap block + collected body into our FAQ structure. The wrapper gets a
      // predictable id="faq-{slug}" so a reader can copy a deep link to this
      // FAQ (the copy-link button below); consumeInitialHash() honours it on
      // arrival.
      const faqSlug = this.slug(h4.textContent || '');
      const wrapper = document.createElement('div');
      wrapper.className = 'hb-faq';
      if (faqSlug) {
        wrapper.id = 'faq-' + faqSlug;
        wrapper.setAttribute('data-faq-id', faqSlug);
      }

      h4.classList.add('hb-faq-summary');
      h4.setAttribute('role', 'button');
      h4.setAttribute('tabindex', '0');
      h4.setAttribute('aria-expanded', 'false');

      // "Copy link" button — copies the deep link for THIS FAQ. Sits at the
      // right end of the H4 (the summary is display:flex, so margin-left:auto
      // floats it right). Only added when we have a slug to link to.
      if (faqSlug) {
        const copyLinkBtn = document.createElement('button');
        copyLinkBtn.type = 'button';
        copyLinkBtn.className = 'hb-faq-copy-link';
        copyLinkBtn.setAttribute('data-hb', 'faq-copy-link');
        copyLinkBtn.setAttribute('data-faq-id', faqSlug);
        copyLinkBtn.setAttribute('aria-label', 'Copy link to this question');
        copyLinkBtn.setAttribute('title', 'Copy link to this question');
        copyLinkBtn.innerHTML = '<span class="hb-faq-copy-link-icon" aria-hidden="true">🔗</span>';
        h4.appendChild(copyLinkBtn);
      }

      const bodyDiv = document.createElement('div');
      bodyDiv.className = 'hb-faq-body';
      for (const el of body) bodyDiv.appendChild(el);

      block.parentNode.insertBefore(wrapper, block);
      wrapper.appendChild(block);
      wrapper.appendChild(bodyDiv);

      const toggle = (e?: Event): void => {
        // Clicks on the copy-link button bubble up through the H4 — bail so its
        // own handler runs alone without also toggling the FAQ.
        if (e) {
          const t = e.target as HTMLElement | null;
          if (t && t.closest('[data-hb="faq-copy-link"]')) return;
          // Prevent SP's auto-injected permalink <a> inside the H4 from
          // hijacking the click.
          e.preventDefault();
        }
        const isOpen = wrapper.classList.toggle('open');
        h4.setAttribute('aria-expanded', String(isOpen));
      };
      h4.addEventListener('click', toggle);
      h4.addEventListener('keydown', (e: KeyboardEvent) => {
        if (e.key === 'Enter' || e.key === ' ') {
          const t = e.target as HTMLElement | null;
          if (t && t.closest('[data-hb="faq-copy-link"]')) return;
          e.preventDefault();
          toggle();
        }
      });
    }

    if (wasObserving) {
      const target = document.querySelector<HTMLElement>(
        '[data-automation-id="mainScrollRegionInnerContent"], [data-automation-id="contentScrollRegion"]'
      );
      if (target) this.attachContentObserver(target);
    }

    // A copied/bookmarked #faq-... link may have arrived before the FAQs were
    // wrapped; now that they exist, try to honour it (runs once).
    this.consumeInitialHash();
  }

  // Copy the deep link to this FAQ so a reader can share or bookmark it. URL is
  // origin + pathname (no query string) + the FAQ hash. Brief icon/title swap
  // gives feedback; falls back to a hidden textarea if the modern clipboard API
  // is unavailable.
  private copyFaqLink(btn: HTMLElement): void {
    const slug = btn.getAttribute('data-faq-id') || '';
    if (!slug) return;
    const url = `${window.location.origin}${window.location.pathname}#faq-${slug}`;

    const flashCopied = (): void => {
      btn.classList.add('copied');
      const icon = btn.querySelector<HTMLElement>('.hb-faq-copy-link-icon');
      const prevIcon = icon ? icon.textContent : '';
      const prevTitle = btn.getAttribute('title');
      if (icon) icon.textContent = '✓';
      btn.setAttribute('title', 'Link copied!');
      window.setTimeout(() => {
        btn.classList.remove('copied');
        if (icon && prevIcon !== null) icon.textContent = prevIcon;
        if (prevTitle) btn.setAttribute('title', prevTitle);
      }, 1400);
    };

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(url).then(flashCopied).catch(() => {
        this.fallbackCopyToClipboard(url);
        flashCopied();
      });
      return;
    }
    this.fallbackCopyToClipboard(url);
    flashCopied();
  }

  // Legacy clipboard write using a hidden textarea + execCommand('copy').
  private fallbackCopyToClipboard(text: string): void {
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.top = '-9999px';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
    } catch { /* best-effort */ }
  }

  // If the page was opened with #faq-... in the URL (a shared deep link),
  // scroll SP's custom scroll container to that FAQ and auto-open it. Native
  // hash navigation doesn't work in SP (scroll happens in a custom region, not
  // the window). Runs once per load — guarded by initialHashConsumed.
  private consumeInitialHash(): void {
    if (this.initialHashConsumed) return;
    const hash = window.location.hash;
    if (!hash || hash.length < 2) {
      this.initialHashConsumed = true;
      return;
    }
    const rawId = hash.substring(1);
    let target: HTMLElement | null = document.getElementById(rawId);
    if (!target) target = document.getElementById('faq-' + rawId);
    if (!target) {
      try {
        target = document.querySelector<HTMLElement>(
          '.hb-faq[data-faq-id="' + (window.CSS && CSS.escape ? CSS.escape(rawId) : rawId) + '"]'
        );
      } catch { /* invalid selector — bail */ }
    }
    if (!target) return; // not in the DOM yet — a later rebuild retries
    this.initialHashConsumed = true;
    window.setTimeout(() => {
      target!.scrollIntoView({ behavior: 'smooth', block: 'start' });
      const faq = target!.closest<HTMLElement>('.hb-faq')
        || target!.querySelector<HTMLElement>('.hb-faq');
      if (faq && !faq.classList.contains('open')) {
        faq.classList.add('open');
        const summary = faq.querySelector<HTMLElement>('.hb-faq-summary');
        if (summary) summary.setAttribute('aria-expanded', 'true');
      }
    }, 250);
  }

  // Discover sections from SP's live DOM. For each CanvasZone, walk every
  // H1/H2/H3 heading inside (in document order) and emit a sidebar entry
  // tagged with its level. Three visual tiers — H1 (group/header), H2 (main),
  // H3 (sub) — match the prototype's nav hierarchy.
  //
  // We set the heading's `id` attribute directly so anchor-link scrolling
  // works natively without needing extra <a> tags.
  private buildSidebar(): void {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;

    const zones = Array.from(document.querySelectorAll<HTMLElement>(
      '[data-automation-id="CanvasZone"]'
    ));

    const seenIds = new Set<string>();
    const items: { id: string; label: string; level: number; el: HTMLElement }[] = [];

    for (const zone of zones) {
      const headings = Array.from(zone.querySelectorAll<HTMLElement>('h1, h2, h3'));
      for (const heading of headings) {
        const label = (heading.textContent || '').trim().replace(/\s+/g, ' ');
        if (!label) continue;

        const level = parseInt(heading.tagName.substring(1), 10);

        const baseId = this.slug(label);
        let id = baseId || `section-${items.length}`;
        let suffix = 2;
        while (seenIds.has(id)) {
          id = `${baseId}-${suffix++}`;
        }
        seenIds.add(id);

        items.push({ id, label, level, el: heading });
      }
    }

    // GROW-ONLY guard. SharePoint lazy-loads and virtualizes this long page
    // (16 zones / 118 FAQs), so any single rebuild pass may see only a fraction
    // of the sections that exist — and the rebuild fires on every scroll-driven
    // DOM mutation. Without this guard, a partial pass would overwrite the full
    // menu with a shorter one (or wipe it to "No sections found"), which is the
    // empty-menu / "Section 1 of 1" bug. So: never let a pass shrink a menu
    // we've already built fuller — only replace it when this pass found at
    // least as many sections. The menu fills in as content loads and never
    // collapses when SP unloads off-screen zones.
    const currentCount = sidebar.querySelectorAll('a.nav-link').length;

    if (items.length === 0) {
      if (currentCount > 0) return; // keep the fuller menu we already have
      sidebar.innerHTML = `<div class="nav-section-label">Sections</div>
        <div style="padding:0.5rem 1.25rem;font-size:0.78rem;opacity:0.7;">No sections found on this page yet.</div>`;
      return;
    }
    if (items.length < currentCount) return; // partial pass — don't shrink

    // Commit. Assign anchor ids only now (on the winning, fullest pass) so the
    // hrefs and heading ids stay consistent with what's rendered.
    for (const it of items) it.el.id = it.id;

    sidebar.innerHTML = items.map(it =>
      `<a href="#${it.id}" class="nav-link nav-h${it.level}">${this.escapeHtml(it.label)}</a>`
    ).join('');
  }

  private slug(s: string): string {
    return String(s).toLowerCase().trim()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 60);
  }

  // ===== Global click delegation =====
  private wireGlobalClicks(): void {
    document.body.addEventListener('click', (e) => {
      const target = e.target as HTMLElement | null;
      if (!target) return;

      const hb = target.closest<HTMLElement>('[data-hb]');
      if (hb) this.handleAction(hb.dataset.hb || '', hb);

      if (!target.closest('.aa-wrap')) {
        document.getElementById('aa-popover')?.classList.remove('open');
      }
      if (!target.closest('.search-wrap')) {
        document.getElementById('search-results')?.classList.remove('open');
      }
    });
  }

  private handleAction(action: string, el: HTMLElement): void {
    switch (action) {
      case 'menu-toggle':
        this.toggleNav();
        return;
      case 'overlay':
        document.getElementById('sidebar')?.classList.remove('open');
        document.getElementById('overlay')?.classList.remove('active');
        return;
      case 'aa-toggle':
        document.getElementById('aa-popover')?.classList.toggle('open');
        return;
      case 'search-toggle':
        this.toggleSearch();
        return;
      case 'font-smaller':
        this.adjustFontSize(-1); return;
      case 'font-larger':
        this.adjustFontSize(1); return;
      case 'font-reset':
        this.adjustFontSize(0); return;
      case 'font': {
        const font = el.getAttribute('data-font') || 'sans';
        this.applyFont(font, true);
        return;
      }
      case 'theme': {
        const theme = el.getAttribute('data-theme') || 'light';
        this.applyTheme(theme);
        return;
      }
      case 'focus-toggle':
        // Same action as the hamburger menu — toggles sidebar visibility.
        this.toggleNav();
        return;
      case 'faq-copy-link':
        this.copyFaqLink(el);
        return;
      case 'agents-toggle': {
        // In embedded mode (Teams Website/SharePoint tab, any iframe host)
        // SP appends ?env=Embedded and hides its suite bar — including the
        // #SUITENAV_COPILOT button we click to open Copilot. So when we're
        // embedded, open the full handbook page in a new tab where the
        // suite bar exists and Copilot works normally.
        if (this.isEmbeddedContext()) {
          const fullUrl = `${window.location.origin}${window.location.pathname}`;
          window.open(fullUrl, '_blank', 'noopener,noreferrer');
          return;
        }
        // SP's Copilot/Agents button lives in the suite bar with
        // id="SUITENAV_COPILOT". A programmatic click opens its panel. SP
        // lazy-loads Copilot on first click, so we show a loading state on
        // our buttons until SP's button reports aria-expanded="true".
        const spBtn = document.getElementById('SUITENAV_COPILOT');
        if (!spBtn) {
          console.warn('[hb] SharePoint Agents button not found on this page');
          return;
        }
        this.setAgentsLoading(true);
        spBtn.click();
        // Fire-and-forget — waitForAgentsOpen drives a UI state change only.
        // Attach .catch to satisfy the no-floating-promises lint rule.
        this.waitForAgentsOpen(spBtn, 6000)
          .then(() => this.setAgentsLoading(false))
          .catch(() => this.setAgentsLoading(false));
        return;
      }
    }
  }

  // True when the page is loaded inside an iframe — typically a Teams tab,
  // Outlook actionable, or a deliberate <iframe> embed. SP also signals this
  // by appending ?env=Embedded to the URL and hiding chrome (suite bar, site
  // header, left rail). We use the OR of both checks for maximum reliability.
  private isEmbeddedContext(): boolean {
    try {
      const env = new URLSearchParams(window.location.search).get('env');
      if (env && env.toLowerCase() === 'embedded') return true;
    } catch {
      // URLSearchParams unsupported / malformed URL — fall through to frame check.
    }
    try {
      return window.self !== window.top;
    } catch {
      // Cross-origin frame access throws; that itself is evidence of embedding.
      return true;
    }
  }

  // Toggle the visual loading state on both Agents entry points.
  private setAgentsLoading(loading: boolean): void {
    const fab = document.getElementById('hb-agents-fab');
    const tbBtn = document.getElementById('btn-agents');
    fab?.classList.toggle('hb-agents-loading', loading);
    tbBtn?.classList.toggle('hb-agents-loading', loading);
  }

  // Poll SP's Copilot button's aria-expanded attribute. Resolves when the
  // panel reports open (aria-expanded="true") or after timeoutMs as a safety.
  private waitForAgentsOpen(spBtn: HTMLElement, timeoutMs: number): Promise<void> {
    return new Promise((resolve) => {
      const start = performance.now();
      const tick = (): void => {
        if (spBtn.getAttribute('aria-expanded') === 'true') {
          resolve();
          return;
        }
        if (performance.now() - start > timeoutMs) {
          resolve();
          return;
        }
        window.requestAnimationFrame(tick);
      };
      tick();
    });
  }

  private applyTheme(theme: string): void {
    document.documentElement.setAttribute('data-theme', theme);
    this.setPref('theme', theme);
    this.markActiveTheme(theme);
  }

  private markActiveTheme(theme: string): void {
    document.querySelectorAll('.theme-swatch').forEach(s => s.classList.remove('active'));
    document.getElementById('sw-' + theme)?.classList.add('active');
  }

  private adjustFontSize(direction: number): void {
    if (direction === 0) this.currentScale = 1;
    else this.currentScale = Math.max(0.75, Math.min(1.5, +(this.currentScale + direction * 0.1).toFixed(2)));
    document.documentElement.style.setProperty('--font-scale', String(this.currentScale));
    this.setPref('fontScale', String(this.currentScale));
    this.updateSizeLabel();
  }

  private updateSizeLabel(): void {
    const el = document.getElementById('size-label');
    if (el) el.textContent = Math.round(this.currentScale * 100) + '%';
  }

  private applyFont(font: string, persist: boolean): void {
    const family = FONTS[font] || FONTS.sans;
    document.documentElement.style.setProperty('--body-font', family);
    document.documentElement.style.setProperty('--heading-font', family);
    if (persist) this.setPref('font', font);
    this.markActiveFont(font);
  }

  private markActiveFont(font: string): void {
    document.querySelectorAll('[id^="btn-font-"]').forEach(b => b.classList.remove('active'));
    document.getElementById('btn-font-' + font)?.classList.add('active');
  }

  // Toggle the sidebar visibility. Behavior depends on viewport:
  //  - Mobile (<=900px): slide overlay in/out, show backdrop
  //  - Desktop (>900px): collapse/expand the always-visible sidebar (push content)
  private toggleNav(): void {
    const sidebar = document.getElementById('sidebar');
    const overlay = document.getElementById('overlay');
    if (!sidebar) return;

    const isMobile = window.innerWidth <= 900;
    if (isMobile) {
      const isOpen = sidebar.classList.toggle('open');
      overlay?.classList.toggle('active', isOpen);
    } else {
      document.documentElement.classList.toggle('hb-nav-collapsed');
    }
  }

  // ===== Sidebar =====
  // Intercept nav-link clicks so we can scroll the actual SP scroll container
  // (window-level anchor navigation doesn't work inside SP's custom scroller).
  private wireSidebar(): void {
    const sidebar = document.getElementById('sidebar');
    if (!sidebar) return;
    sidebar.addEventListener('click', (e) => {
      const link = (e.target as HTMLElement).closest<HTMLAnchorElement>('a.nav-link');
      if (!link) return;

      const href = link.getAttribute('href');
      if (href && href.charAt(0) === '#' && href.length > 1) {
        const target = document.getElementById(href.substring(1));
        if (target) {
          e.preventDefault();
          target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
      }

      if (window.innerWidth <= 900) {
        sidebar.classList.remove('open');
        document.getElementById('overlay')?.classList.remove('active');
      }
    });
  }

  // ===== Scroll spy + reading progress =====
  // Re-runnable so it can refresh whenever the sidebar rebuilds.
  // Modern SP pages scroll inside [data-automation-id="contentScrollRegion"],
  // NOT on the window, so we attach our scroll listener and IntersectionObserver
  // to that container (falling back to window if it's not found).
  private setupScrollSpy(): void {
    if (this.scrollSpyObserver) {
      this.scrollSpyObserver.disconnect();
      this.scrollSpyObserver = null;
    }

    // Remove any prior scroll listener from the previous container or window.
    if (this.scrollUpdateFn) {
      const prevTarget: EventTarget = this.scrollContainer || window;
      prevTarget.removeEventListener('scroll', this.scrollUpdateFn);
      this.scrollUpdateFn = null;
    }

    // Find SP's actual scroll container.
    this.scrollContainer = document.querySelector<HTMLElement>(
      '[data-automation-id="contentScrollRegion"], [data-is-scrollable="true"]'
    );

    const links = Array.from(document.querySelectorAll<HTMLAnchorElement>('.sidebar .nav-link'));
    const targets: { el: HTMLElement; link: HTMLAnchorElement; index: number }[] = [];
    links.forEach((l, index) => {
      const href = l.getAttribute('href');
      if (href && href.charAt(0) === '#' && href.length > 1) {
        const el = document.getElementById(href.substring(1));
        if (el) targets.push({ el, link: l, index });
      }
    });

    if (targets.length > 0 && 'IntersectionObserver' in window) {
      this.scrollSpyObserver = new IntersectionObserver(entries => {
        for (const e of entries) {
          if (e.isIntersecting) {
            links.forEach(l => l.classList.remove('active'));
            const m = targets.find(t => t.el === e.target);
            if (m) m.link.classList.add('active');
          }
        }
      }, {
        root: this.scrollContainer,
        rootMargin: '-80px 0px -60% 0px',
        threshold: 0
      });
      targets.forEach(t => this.scrollSpyObserver!.observe(t.el));
    }

    // Total reflects the FULL menu (which is grow-only and stable), not just
    // the sections SharePoint currently has rendered — otherwise "of N" would
    // shrink as you scroll and SP virtualizes off-screen content.
    const sectionNames = links.map(l => l.textContent || '');
    const rpTotal = document.getElementById('rp-total');
    if (rpTotal) rpTotal.textContent = String(links.length || 1);

    const container = this.scrollContainer;
    const update = (): void => {
      // Progress relative to SP's scroll container (or window if not found).
      const scrollY = container ? container.scrollTop : window.scrollY;
      const scrollH = container
        ? container.scrollHeight - container.clientHeight
        : document.documentElement.scrollHeight - window.innerHeight;
      const pct = scrollH > 0 ? Math.min(100, (scrollY / scrollH) * 100) : 0;
      const pctRound = Math.round(pct);

      const fill = document.getElementById('progress-fill');
      if (fill) fill.style.width = pct.toFixed(1) + '%';

      // "Current" section: last whose top is above 40% of the visible scroll area.
      const containerRect = container ? container.getBoundingClientRect() : null;
      const containerTop = containerRect ? containerRect.top : 0;
      const containerHeight = container ? container.clientHeight : window.innerHeight;
      const triggerLine = containerTop + containerHeight * 0.4;

      // Current section = the last currently-rendered target whose top is above
      // the trigger line; report its position within the FULL menu list so the
      // "Section X of N" stays consistent even when SP has virtualized some
      // sections out of the DOM.
      let currentIndex = 0;
      targets.forEach(t => {
        if (t.el.getBoundingClientRect().top < triggerLine) currentIndex = t.index;
      });
      const rpSection = document.getElementById('rp-section');
      if (rpSection) rpSection.textContent = sectionNames[currentIndex] || '';
      const rpNum = document.getElementById('rp-num');
      if (rpNum) rpNum.textContent = String(currentIndex + 1);
      const rpBarFill = document.getElementById('rp-bar-fill');
      if (rpBarFill) rpBarFill.style.width = pctRound + '%';
      const rpPct = document.getElementById('rp-pct');
      if (rpPct) rpPct.textContent = pctRound + '%';
    };

    const scrollTarget: EventTarget = container || window;
    scrollTarget.addEventListener('scroll', update, { passive: true });
    this.scrollUpdateFn = update;
    update();
  }

  // ===== Search =====
  private wireSearchInput(): void {
    const input = document.getElementById('search-input') as HTMLInputElement | null;
    const results = document.getElementById('search-results');
    if (!input || !results) return;

    input.addEventListener('input', () => {
      const q = input.value.trim().toLowerCase();
      this.clearHighlights();
      if (q.length < 2) { results.classList.remove('open'); return; }
      if (!this.searchIndex) this.searchIndex = this.buildSearchIndex();

      const matches: ISearchHit[] = [];
      for (const item of this.searchIndex) {
        const titleLow = item.title.toLowerCase();
        const textLow = item.text.toLowerCase();
        let score = 0;
        if (titleLow.indexOf(q) >= 0) score += 10;
        if (textLow.indexOf(q) >= 0) score += 1;
        if (score > 0) matches.push({ ...item, score });
      }
      matches.sort((a, b) => b.score - a.score);

      if (matches.length === 0) {
        results.innerHTML = `<div class="search-no-results">No results for "${q.replace(/</g, '&lt;')}"</div>`;
        results.classList.add('open');
        return;
      }

      const shown = matches.slice(0, 8);
      const header = `<div class="search-results-header">${matches.length} result${matches.length !== 1 ? 's' : ''}</div>`;
      const rows = shown.map((m, i) => {
        const textLow = m.text.toLowerCase();
        const idx = textLow.indexOf(q);
        let snippet = '';
        if (idx >= 0) {
          const start = Math.max(0, idx - 40);
          const end = Math.min(m.text.length, idx + q.length + 60);
          snippet =
            (start > 0 ? '...' : '') +
            this.escapeHtml(m.text.substring(start, idx)) +
            '<mark>' + this.escapeHtml(m.text.substring(idx, idx + q.length)) + '</mark>' +
            this.escapeHtml(m.text.substring(idx + q.length, end)) +
            (end < m.text.length ? '...' : '');
        } else {
          snippet = this.escapeHtml(m.text.substring(0, 100)) + (m.text.length > 100 ? '...' : '');
        }
        const icon = m.type === 'glossary' ? '&#128218; ' : m.type === 'question' ? '&#10067; ' : '&#128196; ';
        return `<a class="search-result" data-hb-result-id="${i}">
          <div class="search-result-title">${icon}${this.escapeHtml(m.title)}</div>
          <div class="search-result-snippet">${snippet}</div>
        </a>`;
      }).join('');
      results.innerHTML = header + rows;
      results.classList.add('open');

      results.querySelectorAll<HTMLElement>('.search-result').forEach((el) => {
        el.addEventListener('click', () => {
          const idStr = el.getAttribute('data-hb-result-id') || '0';
          const m = shown[parseInt(idStr, 10)];
          if (!m) return;
          // Open enclosing FAQ collapsible if the target is inside one
          const faq = m.el.closest<HTMLElement>('.hb-faq');
          if (faq && !faq.classList.contains('open')) {
            faq.classList.add('open');
            const summary = faq.querySelector<HTMLElement>('.hb-faq-summary');
            summary?.setAttribute('aria-expanded', 'true');
          }
          // Legacy <details> support (in case any prototype-style collapsibles survived seeding)
          if (m.el.tagName === 'DETAILS') (m.el as HTMLDetailsElement).open = true;
          m.el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          results.classList.remove('open');
          input.classList.remove('open');
          input.value = '';
        });
      });
    });
  }

  private toggleSearch(): void {
    const input = document.getElementById('search-input') as HTMLInputElement | null;
    const results = document.getElementById('search-results');
    if (!input || !results) return;
    const isOpen = input.classList.contains('open');
    if (isOpen) {
      input.classList.remove('open');
      results.classList.remove('open');
      this.clearHighlights();
      input.value = '';
    } else {
      input.classList.add('open');
      window.setTimeout(() => input.focus(), 260);
      if (!this.searchIndex) this.searchIndex = this.buildSearchIndex();
    }
  }

  // Build a search index of everything readers might want to find:
  //   - Section headings (H1/H2/H3) + nearby text → type: section
  //   - FAQ collapsibles (.hb-faq) → type: question (clicking auto-opens)
  //   - Glossary terms (.term[data-def]) → type: glossary (deduplicated)
  // Limited to content inside SP's text web parts to avoid indexing chrome.
  private buildSearchIndex(): ISearchHit[] {
    const idx: ISearchHit[] = [];
    const scope = document.querySelector('[data-automation-id="mainScrollRegionInnerContent"]') || document.body;

    // Section entries: every H1/H2/H3 in content + a snippet of its siblings
    scope.querySelectorAll<HTMLElement>(
      '[data-automation-id="textBox"] h1, [data-automation-id="textBox"] h2, [data-automation-id="textBox"] h3, ' +
      '.ck-content h1, .ck-content h2, .ck-content h3'
    ).forEach(h => {
      const title = (h.textContent || '').trim().replace(/\s+/g, ' ');
      if (!title) return;
      const paras: string[] = [];
      let next = h.nextElementSibling;
      while (next && !/^H[1-6]$/.test(next.tagName) && paras.join(' ').length < 600) {
        const t = (next as HTMLElement).innerText?.trim();
        if (t) paras.push(t);
        next = next.nextElementSibling;
      }
      idx.push({ title, text: paras.join(' '), el: h, type: 'section', score: 0 });
    });

    // FAQ collapsibles
    scope.querySelectorAll<HTMLElement>('.hb-faq').forEach(faq => {
      const summary = faq.querySelector<HTMLElement>('.hb-faq-summary');
      const body = faq.querySelector<HTMLElement>('.hb-faq-body');
      if (!summary) return;
      const title = (summary.innerText || '').trim().replace(/\s+/g, ' ');
      const text = (body?.innerText || '').trim().replace(/\s+/g, ' ').substring(0, 600);
      if (!title) return;
      idx.push({ title, text, el: summary, type: 'question', score: 0 });
    });

    // Glossary terms — dedupe by term text (case-insensitive)
    const seenGloss = new Set<string>();
    scope.querySelectorAll<HTMLElement>('.term[data-def]').forEach(termEl => {
      const title = (termEl.textContent || '').trim().replace(/\s+/g, ' ');
      const text = termEl.getAttribute('data-def') || '';
      if (!title || !text) return;
      const key = title.toLowerCase();
      if (seenGloss.has(key)) return;
      seenGloss.add(key);
      idx.push({ title, text, el: termEl, type: 'glossary', score: 0 });
    });

    return idx;
  }

  private clearHighlights(): void {
    for (const el of this.prevHighlights) {
      const parent = el.parentNode;
      if (parent) {
        parent.replaceChild(document.createTextNode(el.textContent || ''), el);
        parent.normalize();
      }
    }
    this.prevHighlights = [];
  }

  private escapeHtml(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // ===== Keyboard =====
  private wireKeyboardShortcuts(): void {
    document.addEventListener('keydown', (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key === 'k') {
        e.preventDefault();
        this.toggleSearch();
      }
      if (e.key === 'Escape') {
        const input = document.getElementById('search-input') as HTMLInputElement | null;
        if (input?.classList.contains('open')) this.toggleSearch();
        document.getElementById('aa-popover')?.classList.remove('open');
        // Close mobile overlay only; leave desktop sidebar state alone.
        if (window.innerWidth <= 900) {
          document.getElementById('sidebar')?.classList.remove('open');
          document.getElementById('overlay')?.classList.remove('active');
        }
      }
    });
  }
}
