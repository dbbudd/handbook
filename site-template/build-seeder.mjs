#!/usr/bin/env node
/**
 * Builds a browser-pasteable JavaScript blob that populates Home1.aspx with
 * the handbook content from index.html, using your authenticated SharePoint
 * session (cookies) and the SharePoint REST API.
 *
 * Usage:
 *   node build-seeder.mjs
 *
 * Output: site-template/seed-content.generated.js
 *
 * Then:
 *   1. Open Home1.aspx in your browser (any mode, logged in)
 *   2. Open DevTools → Console
 *   3. Paste the contents of seed-content.generated.js
 *   4. Confirm the prompt to overwrite the page content
 *
 * The script uses /_api/sitepages/pages/.../SavePage with a request digest,
 * so the same Site Owner permission that lets you edit the page in the UI
 * is enough — no Entra app registration needed.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import * as cheerio from 'cheerio';

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = resolve(__dirname, '..');
const SOURCE_HTML = resolve(PROJECT_ROOT, 'index.html');
const OUTPUT_JS = resolve(__dirname, 'seed-content.generated.js');

const TARGET_PAGE = 'Home1.aspx';

// ===== Parse index.html into sections + collapsibles =====
function slug(s) {
  return String(s).toLowerCase().trim()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

function parseSection($, el) {
  const $sec = $(el);
  const id = $sec.attr('id');
  const $clone = $sec.clone();

  $clone.find('.section-banner').remove();

  const $titleEl = $clone.find('h2:not(.banner-title), h3').first();
  const title = $titleEl.text().trim().replace(/\s+/g, ' ') || id;
  $titleEl.remove();

  const collapsibles = [];
  $clone.find('details').each((_, d) => {
    const $d = $(d);
    const summaryText = $d.find('summary').first().text().trim().replace(/\s+/g, ' ');
    const $body = $d.clone();
    $body.find('summary').first().remove();
    collapsibles.push({
      title: summaryText || 'Untitled',
      slug: slug(summaryText || id),
      bodyHtml: $body.html()?.trim() ?? ''
    });
    $d.remove();
  });

  const introHtml = $clone.html()?.trim() ?? '';
  return { id, title, introHtml, collapsibles };
}

console.log(`Reading ${SOURCE_HTML}`);
const html = readFileSync(SOURCE_HTML, 'utf8');
const $ = cheerio.load(html);
const sections = $('section[id]').map((_, el) => parseSection($, el)).get();

const totalCollapsibles = sections.reduce((n, s) => n + s.collapsibles.length, 0);
console.log(`Parsed ${sections.length} top-level sections + ${totalCollapsibles} collapsibles`);

// ===== Build canvasContent1 JSON =====
const canvasControls = [];
let zoneIndex = 1;

const anchorHtml = (id) =>
  `<p style="margin:0;padding:0;height:0;overflow:hidden;"><a id="${id}" class="hb-anchor" aria-hidden="true"></a></p>`;

const titleHtml = (text, level = 2) =>
  `<h${level}>${escapeHtml(text)}</h${level}>`;

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

for (const section of sections) {
  // 1) Intro (non-collapsible) — section title + intro HTML
  const introInner = [
    anchorHtml(section.id),
    titleHtml(section.title, 2),
    section.introHtml || ''
  ].join('\n');

  canvasControls.push({
    controlType: 4,
    id: randomUUID(),
    position: {
      zoneIndex,
      sectionIndex: 1,
      controlIndex: 1,
      sectionFactor: 12,
      layoutIndex: 1
    },
    emphasis: {},
    displayMode: 2,
    addedFromPersistedData: true,
    innerHTML: introInner,
    editorType: 'CKEditor'
  });
  zoneIndex++;

  // 2) Each <details> becomes its own collapsible zone
  for (const c of section.collapsibles) {
    // a) Zone descriptor with zoneGroupMetadata = collapsible
    canvasControls.push({
      controlType: 0,
      position: {
        zoneIndex,
        sectionIndex: 1,
        controlIndex: 1,
        sectionFactor: 12,
        layoutIndex: 1
      },
      emphasis: { zoneEmphasis: 0 },
      zoneGroupMetadata: {
        type: 1,
        displayName: c.title,
        isExpanded: false,
        iconAlignment: 'left',
        showDividerLine: true
      },
      addedFromPersistedData: true
    });

    // b) Text web part inside the collapsible zone. We include an <h4> for the
    //    question — the customizer transforms every <h4> into a clickable
    //    FAQ collapsible at render time. <h4> deliberately doesn't show in the
    //    sidebar TOC (TOC only includes H1-H3).
    const collapsibleInner = [
      anchorHtml(`${section.id}-${c.slug}`),
      titleHtml(c.title, 4),
      c.bodyHtml || ''
    ].join('\n');

    canvasControls.push({
      controlType: 4,
      id: randomUUID(),
      position: {
        zoneIndex,
        sectionIndex: 1,
        controlIndex: 2,
        sectionFactor: 12,
        layoutIndex: 1
      },
      emphasis: { zoneEmphasis: 0 },
      displayMode: 2,
      addedFromPersistedData: true,
      innerHTML: collapsibleInner,
      editorType: 'CKEditor'
    });

    zoneIndex++;
  }
}

// Page settings (always last)
canvasControls.push({
  controlType: 0,
  pageSettingsSlice: {
    isDefaultDescription: true,
    isDefaultThumbnail: true,
    isSpellCheckEnabled: true,
    globalRichTextStylingVersion: 1,
    rtePageSettings: { contentVersion: 5 },
    isEmailReady: false
  }
});

console.log(`Generated ${canvasControls.length} canvas controls`);

// ===== Generate the browser-pasteable JS =====
const generatedAt = new Date().toISOString();
const canvasJson = JSON.stringify(canvasControls);

const seederJs = `/*
 * HKIS Handbook content seeder.
 *
 * Generated: ${generatedAt}
 * Source:    ../index.html
 * Target:    SitePages/${TARGET_PAGE}
 * Controls:  ${canvasControls.length} (${sections.length} sections + ${totalCollapsibles} collapsibles)
 *
 * HOW TO RUN:
 *   1. Open the handbook page in your browser (logged into SharePoint).
 *   2. Open DevTools → Console.
 *   3. Paste this entire file's contents and press Enter.
 *   4. Confirm the prompt to overwrite the page content.
 *
 * WARNING: This OVERWRITES the existing CanvasContent1 of Home1.aspx.
 * Your existing content on that page will be replaced.
 */
(async function seedHandbook() {
  const TARGET_PAGE = ${JSON.stringify(TARGET_PAGE)};
  const CANVAS_CONTENT = ${canvasJson};

  // Detect the site web URL from the current location.
  const path = window.location.pathname;
  const spIdx = path.toLowerCase().indexOf('/sitepages/');
  if (spIdx < 0) {
    console.error('[seeder] Run this from a SharePoint page (URL must include /SitePages/).');
    return;
  }
  const siteRelative = path.substring(0, spIdx);
  const siteAbsolute = window.location.origin + siteRelative;
  console.log('[seeder] Site:', siteAbsolute);
  console.log('[seeder] Page:', TARGET_PAGE);
  console.log('[seeder] Controls to write:', CANVAS_CONTENT.length);

  const ok = window.confirm(
    'This will OVERWRITE the contents of ' + TARGET_PAGE + ' with ' +
    CANVAS_CONTENT.length + ' canvas controls (' +
    '${sections.length} sections + ${totalCollapsibles} collapsibles).\\n\\n' +
    'Continue?'
  );
  if (!ok) {
    console.log('[seeder] Cancelled.');
    return;
  }

  console.log('[seeder] Fetching request digest...');
  const digestRes = await fetch(siteAbsolute + '/_api/contextinfo', {
    method: 'POST',
    credentials: 'include',
    headers: { 'accept': 'application/json;odata=nometadata' }
  });
  if (!digestRes.ok) {
    console.error('[seeder] Failed to fetch digest:', digestRes.status, await digestRes.text());
    return;
  }
  const digestData = await digestRes.json();
  const digest = digestData.FormDigestValue;
  console.log('[seeder] Got digest.');

  console.log('[seeder] Fetching page info...');
  const pageRes = await fetch(
    siteAbsolute + "/_api/sitepages/pages/GetByUrl('SitePages/" + TARGET_PAGE + "')",
    {
      credentials: 'include',
      headers: { 'accept': 'application/json;odata=nometadata' }
    }
  );
  if (!pageRes.ok) {
    console.error('[seeder] Failed to fetch page:', pageRes.status, await pageRes.text());
    return;
  }
  const pageData = await pageRes.json();
  console.log('[seeder] Page Id:', pageData.Id, 'Title:', pageData.Title);

  // Discard any prior checkout / draft state before we start. This handles
  // the "423 Locked" / "We cannot save your changes because a site member
  // has ended your editing session" scenarios that occur when the page is in
  // a stale locked state from a previous tab or prior failed seed.
  console.log('[seeder] Discarding any existing draft/checkout...');
  const discardRes = await fetch(
    siteAbsolute + '/_api/sitepages/pages(' + pageData.Id + ')/DiscardPage',
    {
      method: 'POST',
      credentials: 'include',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      }
    }
  );
  if (discardRes.ok) {
    console.log('[seeder] Cleared prior state.');
  } else {
    console.log('[seeder] Discard returned', discardRes.status,
      '(no prior state to clear — continuing).');
  }

  // Check out the page so we can write.
  console.log('[seeder] Checking out page...');
  const checkoutRes = await fetch(
    siteAbsolute + '/_api/sitepages/pages(' + pageData.Id + ')/CheckoutPage',
    {
      method: 'POST',
      credentials: 'include',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      }
    }
  );
  if (checkoutRes.ok) {
    console.log('[seeder] Checked out.');
  } else {
    console.warn('[seeder] Checkout returned', checkoutRes.status,
      '\\u2014 continuing anyway.');
  }

  console.log('[seeder] Saving page... (this may take 10-30 seconds for ' +
    CANVAS_CONTENT.length + ' controls)');
  const saveRes = await fetch(
    siteAbsolute + '/_api/sitepages/pages(' + pageData.Id + ')/SavePage',
    {
      method: 'POST',
      credentials: 'include',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest,
        'If-Match': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        CanvasContent1: JSON.stringify(CANVAS_CONTENT),
        Title: pageData.Title || 'Teaching & Learning Handbook'
      })
    }
  );

  if (!saveRes.ok) {
    const errText = await saveRes.text();
    console.error('[seeder] Save FAILED:', saveRes.status);
    console.error('[seeder] Response:', errText);
    return;
  }
  console.log('[seeder] Saved. Publishing...');

  // Publish so a fresh page load shows the saved canvas (not a draft).
  const publishRes = await fetch(
    siteAbsolute + '/_api/sitepages/pages(' + pageData.Id + ')/Publish',
    {
      method: 'POST',
      credentials: 'include',
      headers: {
        'accept': 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest,
        'If-Match': '*'
      }
    }
  );
  if (publishRes.ok) {
    console.log('[seeder] Published.');
  } else {
    console.warn('[seeder] Publish returned', publishRes.status,
      '\\u2014 saved as draft.');
  }
  console.log('[seeder] Reloading page in 2 seconds...');
  setTimeout(function () { window.location.reload(); }, 2000);
})();
`;

writeFileSync(OUTPUT_JS, seederJs, 'utf8');
console.log(`\nWrote ${OUTPUT_JS}`);
console.log(`File size: ${(seederJs.length / 1024).toFixed(1)} KB`);
console.log(`\nNext: open the file, copy all contents, paste into the DevTools Console on Home1.aspx`);
