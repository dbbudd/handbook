# Handbook Site Template

PnP provisioning template that will stand up the handbook as a SharePoint communication site, including pages, navigation, theme, and static assets.

## Folder layout

```
site-template/
├── template.xml      Main PnP provisioning template (entry point)
├── assets/           Static files uploaded to SiteAssets on provisioning (images, css, etc.)
└── pages/            Modern page XML fragments (included from template.xml)
```

## Authoring

Edit `template.xml` in VS Code — the workspace is configured to use the PnP 2022-09 schema, so you'll get IntelliSense, autocompletion, and validation as you type. Hover any `pnp:` element to see its docs.

Useful references:
- PnP Provisioning schema reference: https://github.com/pnp/PnP-Provisioning-Schema
- Schema docs: https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview

## Local validation

There is no native renderer for PnP templates — they describe what to deploy, not how a page looks. You can validate XML well-formedness and schema conformance via the VS Code XML extension (already configured), and once a Microsoft 365 tenant is connected:

```bash
# from the repo root, after `m365 login`
m365 spo project doctor --filePath site-template/template.xml
```

## Previewing static page content

The existing `../index.html` and anything in `assets/` can be previewed with the **Live Server** VS Code extension — right-click the file in the explorer and choose "Open with Live Server".

## Applying the template (later, once a tenant is connected)

```bash
m365 login
m365 spo site add --type CommunicationSite --title "Handbook" --url https://<tenant>.sharepoint.com/sites/handbook
m365 spo site apply --templatePath ./site-template/template.xml --url https://<tenant>.sharepoint.com/sites/handbook
```
