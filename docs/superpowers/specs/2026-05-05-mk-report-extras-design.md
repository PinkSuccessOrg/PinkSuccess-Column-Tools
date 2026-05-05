# Mary Kay Report Extras — Design

## Goal

Extend the eight existing PinkSuccess Column Tools presets so that, after column extraction, each output is also formatted to match the unit director's newsletter-ready styling. The new specifications come from an updated `NL Report Fields.xlsx` that adds an `EXTRAS` column (per-preset formatting rules) and a final "ALL REPORTS" row (universal rules).

## Source of truth

`~/Downloads/NL Report Fields.xlsx`, columns A–G. Columns A–D + F + G are unchanged. Column E (`EXTRAS`) and row 14 (`ALL REPORTS - Avenir Next LT Pro - 12 font and no boarders and no headers`) are new.

The two presets whose extras read `void` (Star Achiever, Teams) are not in scope — they were already excluded in the previous spec, and "void" simply confirms there is nothing to do.

## All Reports — universal formatting

Three rules apply to every preset run (and to the Custom section's Keep & Reorder; see below for the other two Custom buttons):

1. **Font:** name = `Avenir Next LT Pro`, size = `12`.
2. **No borders:** clear all six border types on the output range — `EdgeTop`, `EdgeBottom`, `EdgeLeft`, `EdgeRight`, `InsideHorizontal`, `InsideVertical`.
3. **No headers:** the output starts with the first data row; the header row is *not* written.

Note on Avenir Next LT Pro: the font ships free with macOS but is not installed by default on Windows. The pages landing site will document this; Excel will silently substitute on machines that lack it.

## Per-preset column formatting

Output column indices are zero-based on the preset's `headers` array. The `numberFormat` for currency is `$#,##0.00` everywhere it's used.

| Preset | Output index | Rule |
|---|---|---|
| OT Star | 1–6 (Wholesale through Pearl) | currency, center |
| Ct Sales | 0 (Rank) | left |
| Ct Sales | 2 (YTD Total) | currency |
| Ct Sharing | 2 (Sem Recruiter Comm Earned) | currency |
| B-Day | — | (none) |
| Anniv | 1 (Number Of Years) | left |
| WHSL | 1 (Unit Wholesale) | currency |
| Checks | 1 (%) | center |
| Checks | 2 (Team Commission) | currency |
| New Cons | — | (none) |

## Data model

`presets.ts` gains an optional `columnFormats` field, expressed as a sparse list of rules so the source maps cleanly to the spreadsheet's "Col B-G - Currency / Center" notation:

```ts
export type Alignment = "Left" | "Center" | "Right";

export interface ColumnFormatRule {
  // Zero-based column index, or inclusive [start, end] range.
  columns: number | [number, number];
  numberFormat?: string;
  horizontalAlignment?: Alignment;
}

export interface Preset {
  id: string;
  label: string;
  headers: string[];
  columnFormats?: ColumnFormatRule[];
}
```

Example (OT Star):

```ts
{
  id: "preset-ot-star",
  label: "OT Star",
  headers: ["Name", "Wholesale $ + Team Building", "Sapphire", "Ruby", "Diamond", "Emerald", "Pearl"],
  columnFormats: [
    { columns: [1, 6], numberFormat: "$#,##0.00", horizontalAlignment: "Center" },
  ],
}
```

Universal rules (font, no borders, no headers) are not in `Preset` — they're hard-coded in `excel-ops.ts` as the always-on policy for the `keepColumnsInOrder` operation.

## Pipeline changes

`excel-ops.ts` `keepColumnsInOrder` becomes:

1. Read used range (unchanged).
2. Detect header row (unchanged).
3. Build `outputRows` with **only the data rows** — drop the header row entirely.
4. Clear the used range (unchanged).
5. Write `outputRows` to the target range starting at `A1` (unchanged otherwise; the target range is now smaller by one row).
6. **New:** apply universal font + clear borders to the entire target range.
7. **New:** for each `ColumnFormatRule` in the preset, resolve the index span and apply `numberFormat` + `horizontalAlignment` to that column slice.
8. Sync.

The status message keeps reporting `Kept X/Y column(s)`. If the source had only a header row and no data, `outputRows` is empty — write nothing, status says `Kept X/Y column(s) (no data rows)` and skip the format pass.

The function's signature changes from `(keepHeaders: string[], setStatus)` to `(opts: { headers: string[]; columnFormats?: ColumnFormatRule[] }, setStatus)`. Callers in `runPreset` and `taskpane.ts` (Custom section) update to pass the new shape.

## Custom section behavior

- **Keep & Reorder** — full All Reports treatment: drops the header row, applies font, clears borders. No per-column formatting (the user typed plain headers). This matches preset behavior.
- **Keep (preserve order)** and **Remove** — unchanged. These are in-place column deletions, not range rewrites; applying "no headers" doesn't fit (no row reshape) and aggressive font/border changes would clobber whatever formatting the user had set up. Surgical operations stay surgical.

## File changes

| File | Change |
|---|---|
| `src/taskpane/presets.ts` | Add `Alignment`, `ColumnFormatRule` types and the `columnFormats` field. Populate rules for OT Star, Ct Sales, Ct Sharing, Anniv, WHSL, Checks. |
| `src/taskpane/excel-ops.ts` | New constants (`UNIVERSAL_FONT`, `UNIVERSAL_FONT_SIZE`). Refactor `keepColumnsInOrder` to take options object, drop header row, and apply formatting after write. New helper `applyColumnFormats(sheet, totalRows, rules)`. New helper `clearAllBorders(range)`. |
| `src/taskpane/headers.ts` | Add `expandColumnRule(rule, columnCount)` — pure helper that turns a `ColumnFormatRule['columns']` into an array of indices, clamped to `[0, columnCount)`. Easy to unit-test. |
| `src/taskpane/headers.test.ts` | Tests for `expandColumnRule`. |
| `src/taskpane/taskpane.ts` | Update Custom-section `keepColumnsInOrder` call site to use the new options shape. |
| `public/index.html` | One-line doc note: Windows users may need to install Avenir Next LT Pro for fonts to render correctly. |

## Out of scope

- Per-cell formatting (only column-level rules).
- Row striping, conditional formatting, header styling.
- Custom font fallback. If Avenir Next LT Pro is missing, Excel substitutes — no detection or error.
- Re-applying formatting to existing data without re-running a preset (you re-run the preset).
- All Reports rules on Custom Keep / Remove (intentionally not applied — see above).

## Risks

- **Avenir Next LT Pro absence on Windows.** Mitigated by documenting in the landing page. The output is still functional, just visually different.
- **Number-format on text columns.** If a "currency" rule lands on a column that contains strings (shouldn't happen for the in-scope presets, but defensive), Excel applies the format spec but doesn't coerce the text. Harmless.
- **Column-rule indices off the end.** If a future preset's `columnFormats` references a column past the headers length, `expandColumnRule` clamps to range and the surplus is silently ignored. Tested.
