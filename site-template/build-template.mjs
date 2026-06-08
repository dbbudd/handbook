#!/usr/bin/env node
/**
 * Builds the PnP provisioning template for the handbook from index.html.
 *
 * Usage:
 *   node build-template.mjs                          # structure only, no content
 *   node build-template.mjs --seed-content           # include section content (additive)
 *   node build-template.mjs --seed-content --reset-content
 *       # explicitly authorize overwriting any prior content on the live page
 *
 * Output: site-template/template.generated.xml
 *
 * The generated template targets exactly one page (SitePages/Home1.aspx) on the
 * already-existing Curriculum site. It NEVER touches other pages, lists, or
 * theme settings on that site. Re-running without --reset-content will not
 * overwrite content that authors have edited in SharePoint.
 *
 * Layout strategy: each top-level <section id="..."> in index.html becomes
 *   - one non-collapsible SP section (intro content, with a hidden anchor)
 *   - one collapsible SP section per <details> inside (using <summary> as title)
 * This lets editors work with SharePoint's native collapsibles instead of
 * custom JS, while preserving the prototype's anchor-link navigation.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as cheerio from 'cheerio';

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = resolve(__dirname, '..');
const SOURCE_HTML = resolve(PROJECT_ROOT, 'index.html');
const OUTPUT_XML = resolve(__dirname, 'template.generated.xml');

const TARGET_PAGE_NAME = 'Home1';
const TARGET_PAGE_TITLE = 'Teaching & Learning Handbook';

const args = new Set(process.argv.slice(2));
const SEED_CONTENT = args.has('--seed-content');
const RESET_CONTENT = args.has('--reset-content');

if (RESET_CONTENT && !SEED_CONTENT) {
  console.error('--reset-content has no effect without --seed-content');
  process.exit(1);
}

console.log(`Reading ${SOURCE_HTML}`);
const html = readFileSync(SOURCE_HTML, 'utf8');
const $ = cheerio.load(html);

function slug(s) {
  return String(s).toLowerCase().trim()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

function xmlAttr(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Parse a top-level <section id="..."> into:
 *   { id, title, introHtml, collapsibles: [{ title, slug, bodyHtml }] }
 */
function parseSection(el) {
  const $sec = $(el);
  const id = $sec.attr('id');

  const $clone = $sec.clone();

  // Drop visual section banners — section group headers don't transfer to SP
  $clone.find('.section-banner').remove();

  // Section title = first non-banner H2/H3 inside
  const $titleEl = $clone.find('h2:not(.banner-title), h3').first();
  const title = $titleEl.text().trim().replace(/\s+/g, ' ') || id;
  $titleEl.remove();

  // Extract <details> elements as collapsibles; remove them from the clone
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

const sections = $('section[id]').map((_, el) => parseSection(el)).get();

console.log(`\nDiscovered ${sections.length} top-level sections:`);
for (const s of sections) {
  const collapsibleNote = s.collapsibles.length
    ? ` (+${s.collapsibles.length} collapsibles)`
    : '';
  const introNote = s.introHtml ? `✓` : '∅';
  console.log(`  ${introNote} #${s.id.padEnd(28)} ${s.title}${collapsibleNote}`);
}

// XML builders ---------------------------------------------------------------

const textWebPart = (html, controlId, order = 1, column = 1) => {
  if (!html?.trim()) return '';
  return `          <pnp:CanvasControl WebPartType="Text" ControlId="${xmlAttr(controlId)}" Order="${order}" Column="${column}">
            <pnp:CanvasControlProperties>
              <pnp:CanvasControlProperty Key="Text"><![CDATA[${html}]]></pnp:CanvasControlProperty>
            </pnp:CanvasControlProperties>
          </pnp:CanvasControl>`;
};

const introSection = (s, order) => {
  const anchorHtml = `<a id="${s.id}" class="hb-anchor" aria-hidden="true"></a>`;
  const introWithAnchor = `${anchorHtml}\n${s.introHtml || ''}`;
  return `      <!-- Section anchor: #${s.id} (${s.title}) -->
      <pnp:Section Order="${order}" Type="OneColumn" CanvasEmphasis="None">
        <pnp:Controls>
${SEED_CONTENT ? textWebPart(introWithAnchor, `${s.id}-intro`) : `          <!-- Intro not seeded; rerun with --seed-content -->`}
        </pnp:Controls>
      </pnp:Section>`;
};

const collapsibleSection = (c, parentId, order) => {
  const anchorHtml = `<a id="${parentId}-${c.slug}" class="hb-anchor" aria-hidden="true"></a>`;
  const bodyWithAnchor = `${anchorHtml}\n${c.bodyHtml || ''}`;
  return `      <pnp:Section Order="${order}" Type="OneColumn" CanvasEmphasis="None" Collapsible="true" IsExpanded="false" DisplayName="${xmlAttr(c.title)}" ShowDividerLine="true">
        <pnp:Controls>
${SEED_CONTENT ? textWebPart(bodyWithAnchor, `${parentId}-${c.slug}`) : `          <!-- Collapsible not seeded; rerun with --seed-content -->`}
        </pnp:Controls>
      </pnp:Section>`;
};

const allSectionsXml = (() => {
  let order = 0;
  const blocks = [];
  for (const s of sections) {
    blocks.push(introSection(s, ++order));
    for (const c of s.collapsibles) {
      blocks.push(collapsibleSection(c, s.id, ++order));
    }
  }
  return blocks.join('\n');
})();

const overwriteBehaviorComment = RESET_CONTENT
  ? `  <!-- WARNING: --reset-content was specified. Existing content on Home1.aspx WILL be overwritten on apply. -->`
  : `  <!-- Safe-by-default: this template never overwrites existing content unless --reset-content is passed. -->`;

const totalCollapsibles = sections.reduce((n, s) => n + s.collapsibles.length, 0);
const generatedAt = new Date().toISOString();

const xml = `<?xml version="1.0" encoding="utf-8"?>
<!--
  GENERATED FILE — do not edit by hand.
  Run \`npm run build\` (structure only) or \`npm run build:seed\` (with content) in site-template/.

  Generated:           ${generatedAt}
  Source:              ../index.html
  Target page:         SitePages/${TARGET_PAGE_NAME}.aspx
  Content seeding:     ${SEED_CONTENT ? 'enabled' : 'disabled'}
  Reset content:       ${RESET_CONTENT ? 'YES — will overwrite live content' : 'no'}
  Sections found:      ${sections.length}
  Collapsibles found:  ${totalCollapsibles}
  Total SP sections:   ${sections.length + totalCollapsibles}
-->
<pnp:Provisioning xmlns:pnp="http://schemas.dev.office.com/PnP/2022/09/ProvisioningSchema"
                  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                  xsi:schemaLocation="http://schemas.dev.office.com/PnP/2022/09/ProvisioningSchema https://raw.githubusercontent.com/pnp/PnP-Provisioning-Schema/master/ProvisioningSchema-2022-09-FullSchema.xsd">

  <pnp:Preferences Generator="handbook-build-template" Version="0.2.0" />

${overwriteBehaviorComment}

  <pnp:Templates ID="HANDBOOK-GENERATED">
    <pnp:ProvisioningTemplate ID="HANDBOOK-MAIN" Version="0.2" Scope="RootSite">

      <pnp:ClientSidePages>
        <pnp:ClientSidePage PageName="${TARGET_PAGE_NAME}.aspx"
                            PromoteAsTemplate="false"
                            Title="${xmlAttr(TARGET_PAGE_TITLE)}"
                            Overwrite="${RESET_CONTENT ? 'true' : 'false'}"
                            Layout="Article">
          <pnp:Header Type="None" />
          <pnp:Sections>
${allSectionsXml}
          </pnp:Sections>
        </pnp:ClientSidePage>
      </pnp:ClientSidePages>

    </pnp:ProvisioningTemplate>
  </pnp:Templates>

</pnp:Provisioning>
`;

writeFileSync(OUTPUT_XML, xml, 'utf8');

console.log(`\nWrote ${OUTPUT_XML}`);
console.log(`Mode:                ${SEED_CONTENT ? 'structure + content' : 'structure only'}`);
console.log(`Top-level sections:  ${sections.length}`);
console.log(`Collapsibles:        ${totalCollapsibles}`);
console.log(`Total SP sections:   ${sections.length + totalCollapsibles}`);
if (RESET_CONTENT) {
  console.log('\nWARNING: --reset-content enabled. Applying this template will OVERWRITE live page content.');
}
