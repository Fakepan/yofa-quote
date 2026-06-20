# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

This is a zero-dependency, single-file HTML app. No build step, no package manager.

```bash
# Open directly in browser
open index.html
# Or serve locally (Python)
python3 -m http.server 8080
```

## Architecture

Everything lives in `index.html` — one flat script block (~450 lines). There is no bundler, no framework, no modules.

### Data Layer
- All state is stored in `localStorage` under key `yofa_data_v1`.
- `D` is the global in-memory data object (type matches `DEFAULT_DATA`).
- `load()` / `save()` sync between `D` and `localStorage`.
- Schema: `{ company, customers[], quotes[], seq }`.

### Rendering
- `render()` is a full re-render — it replaces `#app`'s innerHTML.
- `view` is a global string (`"quotes" | "edit" | "board" | "customers" | "settings"`); `go(v)` switches views.
- All event handlers are inline `onclick` / `onchange` attributes calling global functions.
- `rerenderList()` is a partial re-render optimization for the quote list filter.

### Quote Calculation
- `quoteTotals(q)` computes subtotal, tax (5%), total, cost, margin, marginRate from `q.items`.
- Cost fields are internal-only — never appear in the print/PDF output.

### Excel Import
- Uses `xlsx` (SheetJS) from CDN.
- `importExcel()` → `parseSheet()` → `matchHeader()`: fuzzy header matching against `HEADER_SYN` synonyms, then row parsing with keyword-based skip rules.
- Imported items get `cost = unitPrice * 0.7` as a default estimate.

### Print / PDF
- `printQuote()` writes customer-facing HTML into `#printArea` (hidden div), then calls `window.print()`.
- CSS `@media print` hides everything except `#printArea`.
- Cost and margin data are never written into `#printArea`.

### XSS Safety
- All user-generated strings must go through `esc()` before being interpolated into HTML strings.
