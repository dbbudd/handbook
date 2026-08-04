/**
 * Shared load-timing instrumentation for the handbook customizers.
 *
 * WHY THIS EXISTS
 * ---------------
 * On Home1.aspx the bare SharePoint page renders first and the handbook chrome
 * only appears ~20 seconds later. Reading the code rules out our own ordering as
 * the cause: `activate()` applies `.hb-active`, the theme attribute and the
 * toolbar synchronously in its first few statements, so once activate() starts
 * the page looks right within milliseconds. The delay therefore sits somewhere
 * upstream of activate(), and there are only three candidates:
 *
 *   1. SharePoint doesn't ASK for our bundle until late (it defers Application
 *      Customizers behind the page's own critical render path, and this page is
 *      heavy — 16 CanvasZones, 118 FAQs).
 *   2. SharePoint asks early but the bundle DOWNLOADS slowly (ClientSideAssets
 *      served without the Office 365 CDN).
 *   3. The bundle arrives quickly but activate() is held back by the
 *      `document.readyState === 'loading'` -> DOMContentLoaded gate in onInit.
 *
 * Those three have completely different fixes, so guessing is expensive. This
 * module makes one page load answer the question.
 *
 * HOW TO READ THE OUTPUT
 * ----------------------
 * Load Home1.aspx normally (NO ?debugManifestsFile — see CLAUDE.md §9, the debug
 * URL runs the deployed copy and the debug copy at once and would double every
 * mark). About 12 seconds after activation a table prints to the console under
 * `[hb-timing]`. `window.__hbTiming.dump()` reprints it on demand.
 *
 * Every number is milliseconds since navigation start, so marks and the browser's
 * own Navigation/Resource Timing entries are directly comparable. The diagnosis:
 *
 *   - `bundle requested` late (~15s+)      -> cause 1. Page weight. The fix is
 *                                             content-side: fewer CanvasControls.
 *   - `bundle requested` early, long
 *     download duration                    -> cause 2. Enable the Office 365 CDN.
 *   - bundle finished early but
 *     `activate:enter` much later          -> cause 3. Drop the DOMContentLoaded
 *                                             gate; split chrome from wiring.
 *
 * This module is diagnostic scaffolding, not a feature. Once the 20s is
 * understood and fixed, delete it and the `mark()` calls in both customizers.
 */

const STORE_KEY: string = '__hbTiming';

/** A single recorded moment, in ms since navigation start. */
export interface IHbMark {
  /** Which customizer recorded it — 'experience' or 'glossary'. */
  source: string;
  /** Short stable label, e.g. 'activate:enter'. */
  label: string;
  /** performance.now() at the moment of recording. */
  atMs: number;
  /** Optional free text (pathname, retry count, cache hit/miss...). */
  detail?: string;
}

interface IHbTimingStore {
  marks: IHbMark[];
  dumpTimer: number | undefined;
  dump: () => void;
}

/** Milliseconds since navigation start, rounded to whole ms. */
function nowMs(): number {
  return Math.round(performance.now());
}

/**
 * The store lives on `window`, not in module scope, because the two customizers
 * are compiled into SEPARATE bundles — each gets its own private copy of this
 * module. Only a window-level object lets both write into one timeline.
 */
function getStore(): IHbTimingStore {
  const w = window as unknown as Record<string, IHbTimingStore | undefined>;
  let s = w[STORE_KEY];
  if (!s) {
    s = { marks: [], dumpTimer: undefined, dump: dump };
    w[STORE_KEY] = s;
  }
  return s;
}

/** Record a moment. Cheap and exception-safe — never let instrumentation break the page. */
export function mark(source: string, label: string, detail?: string): void {
  try {
    getStore().marks.push({ source: source, label: label, atMs: nowMs(), detail: detail });
  } catch {
    /* instrumentation must never throw */
  }
}

/**
 * Navigation Timing, reduced to the four numbers that matter here. All values are
 * ms since navigation start, matching our marks.
 */
function navigationRows(): IHbMark[] {
  const rows: IHbMark[] = [];
  try {
    const entries = performance.getEntriesByType('navigation');
    if (!entries.length) return rows;
    const nav = entries[0] as unknown as {
      responseEnd: number;
      domContentLoadedEventEnd: number;
      domComplete: number;
      loadEventEnd: number;
    };
    rows.push({ source: 'browser', label: 'html responseEnd', atMs: Math.round(nav.responseEnd) });
    rows.push({ source: 'browser', label: 'DOMContentLoaded', atMs: Math.round(nav.domContentLoadedEventEnd) });
    rows.push({ source: 'browser', label: 'domComplete', atMs: Math.round(nav.domComplete) });
    rows.push({ source: 'browser', label: 'load event', atMs: Math.round(nav.loadEventEnd) });
  } catch {
    /* Navigation Timing L2 unavailable — skip */
  }
  return rows;
}

/**
 * Resource Timing for our own bundles. This is the single most diagnostic
 * measurement available: `startTime` is the moment SharePoint actually ASKED for
 * our code, and `duration` is how long the fetch took. Together they split
 * "SharePoint got to us late" from "the download was slow".
 */
function bundleRows(): IHbMark[] {
  const rows: IHbMark[] = [];
  try {
    const entries = performance.getEntriesByType('resource') as unknown as Array<{
      name: string;
      startTime: number;
      duration: number;
      transferSize?: number;
    }>;
    for (const e of entries) {
      const name = (e.name || '').toLowerCase();
      const isOurs =
        name.indexOf('handbook-experience') !== -1 ||
        name.indexOf('glossary-tooltips') !== -1 ||
        name.indexOf('hkis-handbook') !== -1;
      if (!isOurs) continue;
      const file = e.name.substring(e.name.lastIndexOf('/') + 1).split('?')[0];
      const kb = typeof e.transferSize === 'number' && e.transferSize > 0
        ? `${Math.round(e.transferSize / 1024)}KB `
        : '';
      rows.push({
        source: 'network',
        label: `bundle requested: ${file}`,
        atMs: Math.round(e.startTime),
        detail: `${kb}download took ${Math.round(e.duration)}ms, finished at ${Math.round(e.startTime + e.duration)}ms`
      });
    }
  } catch {
    /* Resource Timing unavailable — skip */
  }
  return rows;
}

/**
 * Print the merged timeline. Safe to call repeatedly; exposed as
 * `window.__hbTiming.dump()` so it can be re-run from the console.
 */
export function dump(): void {
  try {
    const marks = getStore().marks.slice();
    const all = marks.concat(navigationRows()).concat(bundleRows());
    all.sort((a, b) => a.atMs - b.atMs);

    const first = all.length ? all[0].atMs : 0;
    const rows = all.map(m => ({
      'at (ms)': m.atMs,
      '+since first': m.atMs - first,
      source: m.source,
      event: m.label,
      detail: m.detail || ''
    }));

    /* eslint-disable no-console */
    console.log(
      `%c[hb-timing] handbook load timeline — ${all.length} events, ` +
      `all times in ms since navigation start`,
      'font-weight:bold'
    );
    if (typeof console.table === 'function') {
      console.table(rows);
    } else {
      console.log(rows);
    }
    console.log('[hb-timing] re-run with window.__hbTiming.dump() — raw marks in window.__hbTiming.marks');
    /* eslint-enable no-console */
  } catch {
    /* never let the dump break the page */
  }
}

/**
 * Schedule a single dump `delayMs` after the call. Repeated calls push the timer
 * back rather than stacking, so whichever customizer settles last decides when
 * the table prints — meaning the table always contains both customizers' marks.
 */
export function scheduleDump(delayMs: number): void {
  try {
    const s = getStore();
    if (s.dumpTimer !== undefined) window.clearTimeout(s.dumpTimer);
    s.dumpTimer = window.setTimeout(() => {
      s.dumpTimer = undefined;
      dump();
    }, delayMs);
  } catch {
    /* instrumentation must never throw */
  }
}
