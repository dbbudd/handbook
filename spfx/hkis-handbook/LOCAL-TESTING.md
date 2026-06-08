# Local testing — running the customizers against real Home1.aspx

You can see both customizers running on the real `Home1.aspx` page in your browser,
**without deploying anything**, using SPFx's debug mode. The code is served from
`https://localhost:4321/` and loaded into your authenticated SharePoint browser
session via URL query parameters. Nothing gets installed; only your browser
session sees it. Closing the tab ends the experiment.

## One-time setup

Trust the local dev certificate (adds a cert to your macOS Keychain so
`https://localhost:4321/` is trusted by your browser):

```bash
cd spfx/hkis-handbook
gulp trust-dev-cert
```

macOS will prompt for your login password to update Keychain. Approve.

## Each time you want to test

1. **Sign into SharePoint in your browser first.** Open
   <https://hkis.sharepoint.com/sites/Curriculum> in Chrome/Edge and complete
   normal login (with MFA). Stay signed in.

2. **From this folder, start the dev server**:

   ```bash
   cd spfx/hkis-handbook
   gulp serve
   ```

   `gulp serve` will:
   - bundle the customizers
   - start an HTTPS server on `localhost:4321`
   - open a browser tab pointing at
     `https://hkis.sharepoint.com/sites/Curriculum/SitePages/Home1.aspx`
     with a long `?loadSPFX=true&debugManifestsFile=...&customActions=...`
     suffix that tells SharePoint to load our local code.

3. **SharePoint will show a yellow warning bar**:
   *"Allow debug scripts? Debug scripts can pose a security threat."*
   Click **Load debug scripts**. (You can dismiss it per session — it's there
   for safety because real SharePoint pages are loading code from localhost.)

4. The handbook toolbar appears at the top of the page; the sidebar slides in
   from the left; theme switching, font scaling, search, focus mode, and
   glossary tooltips all work.

5. To stop: `Ctrl+C` in the terminal.

## Useful variants

Test just one customizer in isolation:

```bash
gulp serve --config handbookExperience    # toolbar/sidebar/themes/search only
gulp serve --config glossaryTooltips      # just term tooltips
gulp serve --config default               # both (this is what `gulp serve` runs)
```

## What you'll see vs. what's missing

- ✅ The toolbar, sidebar TOC, themes, font scaling, dyslexic font, focus mode,
  reading position bar, search dropdown, keyboard shortcuts (Cmd/Ctrl+K, Esc).
- ✅ Glossary tooltips — but **only if** the page already has elements with
  `<span class="term" data-def="...">` markup. SharePoint's rich text editor
  doesn't expose `class` and `data-` attributes directly, so to test you'll
  either need:
  - some seeded content with `term` spans pasted in via an "Embed" web part
    (Home1.aspx can host an Embed web part with raw HTML), OR
  - manually inject a test span via DevTools to verify the tooltip behavior.
- ⏳ The full 114 sections of seeded content from the build script — not until
  we can actually apply `template.generated.xml` to the site, which is still
  waiting on the m365 CLI Entra app from IT.
- ⏳ Real `.sppkg` deployment — waits on the site collection app catalog being
  enabled.

## Troubleshooting

- **Browser refuses to load scripts**: confirm you trusted the dev cert
  (`gulp trust-dev-cert`). Visit <https://localhost:4321/temp/manifests.js>
  directly — it should load without warnings.
- **"Allow debug scripts" bar doesn't appear**: make sure you didn't open the
  URL in an incognito/private window where the SP session isn't logged in.
- **Customizers don't activate**: check the URL guard. The customizers only run
  on a path ending in `/sitepages/home1.aspx`. If you renamed Home1.aspx, update
  `HANDBOOK_PAGE_PATH` in both customizers' `.ts` files (or pass `pagePath` via
  the `properties` block in `serve.json`).
