# Phase 4 — PowerDashboard Comparison Report

**Date:** 2026-11-06
**Scope:** `src/Modules/modPowerDashboardEngine.bas` (local, authoritative) vs. the copy embedded in `workbook/XLSmart_User.xlam` (extracted from `xl/vbaProject.bin` via oletools 0.60.2).

**Artifacts** (all in `docs/phase4-evidence/`):
- `vba_extracted.txt` — full olevba dump (55,584 lines); module block at lines 12861–13237.
- `xlam_pd_engine.bas` — extracted embedded body (374 lines).
- `pd_engine.diff` — unified diff (700 lines, single hunk `@@ -1,323 +1,374 @@`).
- `vbaProject.bin` — extracted VBA project from `workbook/XLSmart_User.xlam`.
- `olevba_err.txt` — olevba stderr (benign pyparsing warnings only).

## Verdict

The embedded module is an **older, structurally different implementation** — not a byte-identical copy. The local export is authoritative (README.md:55). The embedded version still contains dead p-code residue (`Enabled`/`54582` marker) from a prior module state. **No data-loss or functional regression was found in the local source;** drift is one-directional (local is newer).

## 1. Identity

| | Local | Embedded (xlam) |
|---|---|---|
| SHA-256 (xlam container) | `d95094b2…ff294` | `d95094b2…ff294` (identical file) |
| Module file | `modPowerDashboardEngine.bas` | same name, no `Attribute VB_Name` line |
| Lines | 323 | 374 |
| Entry signature | `CreateDashboard(dashTitle, theme, currencyCode, isMulti, rangeOrSheets, startRow, smartLock, gridlinesShow)` | `CreateDashboard(dashTitle, theme, currencyCodeCode, isMulti, rangeOrSheets, startRow, smartLock)` |

Key signature drift:
- Embedded param typo: `currencyCodeCode` (local: `currencyCode`).
- Embedded **lacks `gridlinesShow`** — the gridlines feature does not exist in the shipped add-in.
- Embedded `BuildDashboardHTML` also drops `gridlinesShow` (config has no `showGrid`).

## 2. Error-code contract

Both use `vbObjectError + 100…103` for workbook/range/data guards with the same semantics (100 = no workbook, 101 = ThisWorkbook guard, 102 = invalid range, 103 = no readable data).

**Local adds `+104`** ("Dashboard HTML file could not be created.") — embedded has no post-write file check. The local version is strictly safer.

Embedded uses `On Error GoTo CleanFail` with a `CleanFail:` label; local uses `On Error GoTo ErrHandler` + `ErrHandler:`. Cosmetic only.

## 3. Multi-sheet path

Functionally equivalent:
- Both split `rangeOrSheets` on `"|"`, require 3 parts, resolve `Workbooks(parts(0))` / `Worksheets(parts(1))` / `Range(parts(2))`.
- Both cap at 50 selected sheets, emit `sheets:[…]` with `name`,`headers`,`rows` (max 50 rows/sheet), `currency`/`currencySym`, `showGrid`, `smartLock`.
- Embedded uses `HeadersFromRange`/`RowsFromRange` + `JoinQuoted`; local uses the newer helper family (`MultiSheetsJSON`, `SheetEntryJSON`, `RangeMatrixJSON`, `HeaderText`).

Formatting differs (embedded uses multi-line `If…Then`, unindented continuations; local is one-line/indented) — **not semantic.**

## 4. Helper inventory

| Embedded-only | Local-only (newer) |
|---|---|
| `HeadersFromRange` | `CurrencyJsLiteral`, `ReplaceBetween`, `TemplateThemeName` |
| `RowsFromRange` | `WriteUtf8`, `MultiSheetsJSON`, `SingleSheetsJSON` |
| `JoinQuoted` | `SheetEntryJSON`, `RangeMatrixJSON`, `HeaderText` |

`JoinQuoted` vs `RangeMatrixJSON`: both emit quoted CSV rows; local's version handles escaping via `JsonValue`/`JsonQuote` equivalently. No behavior gap.

## 5. HTML output & config injection

Both write UTF-8 HTML to `Environ$("TEMP") & "\power_dashboard.html"`, open via `ThisWorkbook.FollowHyperlink` with `Shell "cmd /c start …"` fallback, and `MsgBox "Dashboard created successfully!"`.

Config injection marker identical: `/*--AICONFIG_START--*/ … /*--AICONFIG_END--*/`, JS `let aiConfig = {currency,currencySym,showGrid,smartLock,sheets}` — except embedded omits `showGrid` (see §1).

Local adds `WriteUtf8` (explicit UTF-8 file writer); embedded relies on `Open … For Output` default encoding — a potential non-ASCII risk in the embedded copy only.

## 6. Dead p-code residue (`Enabled` / `54582 bytes`)

`vba_extracted.txt:48096` contains a comment `' VBA/modPowerDashboardEngine - 54582 bytes` followed by `FuncDefn (Public Sub Enabled(ByVal modPowerDashboardEngine As String))`. This is **stale p-code from a previous module state**, not a live duplicate module — no `VBA MACRO` header exists between lines 44000–49000 (headers searched at 38000–47100 ranges). Surrounding lines (47980–48010) are unrelated form p-code (`btnApply_Click`, `ValidationTools_Phone`). **No action required;** it does not affect the shipped `modPowerDashboardEngine` block at 12861–13237.

## 7. Risk assessment

| Item | Risk | Notes |
|---|---|---|
| Missing `gridlinesShow` in xlam | **Medium** | Shipped add-in cannot toggle gridlines; local ribbon/entry must pass the param only when targeting local build. |
| `currencyCodeCode` typo in xlam | Low | Callers in xlam must match the typo; local callers use `currencyCode`. |
| No `+104` file check in xlam | Low | Silent failure if TEMP write fails. |
| Default-encoding write in xlam | Low | Non-ASCII titles/headers may mojibake in embedded copy. |
| Dead `Enabled` p-code | Info | Harmless residue. |

## 8. Recommendation

Treat the local `modPowerDashboardEngine.bas` as the sole source of truth. The xlam-embedded copy is a **pre-gridlines, pre-UTF8-writer snapshot** and should be regenerated from local source during the next packaging step (release phase). No local code changes are required as a result of this comparison.
