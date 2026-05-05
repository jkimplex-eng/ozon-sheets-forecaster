# AGENTS.md

This repository is the Ozon Google Sheets Apps Script project.

## Scope

Work only on:

- `Code.gs`
- `README.md`
- files under `docs/`

Do not modify unrelated repositories. If `Code.gs` is missing, stop immediately and report that the wrong repository was selected.

## Runtime

This is not a Go, Python, Node, or service-account project. It is Google Apps Script copied into Google Sheets.

Do not add:

- `go.mod`
- `go.sum`
- `package.json`
- `service-account.json`
- `.env` with real secrets

## Security

Never commit or print:

- Ozon Seller API key
- Ozon Performance client secret
- access token
- Authorization header value

Use only placeholder variable names in docs.

## Google Sheets flow

The current deployment model is container-bound Apps Script:

1. Open Google Sheet.
2. Extensions -> Apps Script.
3. Paste `Code.gs`.
4. Save.
5. Reload the sheet.
6. Use the `Ozon` menu.

Deploy is not required for the menu-based flow.

## Before editing

Run or inspect the equivalent of:

```bash
ls
rg --files | rg "Code.gs|README.md"
```

If `Code.gs` is missing, stop.

## Commit style

Use focused commits, for example:

- `Add ads debug flow and Summary ad diagnostics`
- `Fix PnL finance ads and stock SKU mapping`
- `Add Ozon Performance statistics reports`
