# Mary Kay Report Presets — Design

## Goal

Replace the three placeholder presets in the PinkSuccess Column Tools task pane with eight buttons sourced from `NL Report Fields.xlsx`. Each button keeps a fixed set of columns from the active sheet, in a fixed order, dropping everything else.

## Source of truth

`NL Report Fields.xlsx` lists the buttons. Example reports in `~/Downloads/mk_report_examples/` define the actual column names that need to match. Where the spreadsheet's "FIELDS NEEDED" wording differs from the actual report column header, the actual header wins (verified against the example file).

Two rows from the spreadsheet are intentionally **out of scope**:

- **Teams** — fields are described structurally ("Col A — names in BOLD"), not as named columns. The note explicitly says "may need to talk about this one."
- **Star Achiever** — would require multi-row-header parsing, previous-quarter date logic, blank-row filtering, and a custom Mary Kay tier comparator. Deferred to a future iteration.

## The eight presets

| Button | Source report | Output columns (in order) |
|---|---|---|
| OT Star | Star Consultant Tracking | Name, Wholesale $ + Team Building, Sapphire, Ruby, Diamond, Emerald, Pearl |
| Ct Sales | Unit Recognition Seminar YTD Court of Sales | Rank, Conslt Name, YTD Total |
| Ct Sharing | Unit Recognition Seminar YTD Court of Sharing | Name, Sem Qual Team Member, Sem Recruiter Comm Earned |
| B-Day | Birthdays, Anniversaries and Addresses (Birthdays variant) | Name, Birth Date |
| Anniv | Birthdays, Anniversaries and Addresses (Anniversaries variant) | Name, Number Of Years |
| WHSL | Unit Recognition Unit Wholesale Scoreboard | Name, Unit Wholesale |
| Checks | Team Commissions | Name, %, Team Commission |
| New Cons | Unit Recognition New Unit Members | New Consultant, From, Recruiter |

Header strings in this table are the canonical preset config, written exactly as they appear in the source files. The matcher is case-insensitive and tolerates one normalization quirk: trailing `*` (Star Consultant Tracking ships `Wholesale $ + Team Building*`, but the preset config carries the cleaner form without the asterisk).

## Multi-row header detection

Three of the source reports (Star Consultant Tracking, Team Commissions, unit stars and consultant consistency) place a "grouping" row above the actual header row. Row 1 typically contains merged-cell labels like `Personal Team` or `Contest Credit Needed`, leaving most of its cells empty after un-merging. Row 2 carries the real headers.

The other reports use a plain single header row at row 1.

**Detection rule.** Count non-empty cells (after `String(cell ?? "").trim()`) in the first two rows of the used range. The row with more non-empty cells is the header row. If that's row 2, row 1 is treated as a grouping row and discarded from the output. If row 1 wins (or only one row exists), there is no grouping row.

This is sufficient for the eight in-scope reports. It is not load-bearing for any preset that needs the grouping row's content — Star Achiever, which would need that, is out of scope.

## Pipeline

1. Read the used range's `values`, `rowCount`, `columnCount`.
2. Detect the header row index `h` (0 or 1) per the rule above. Data starts at row `h + 1`.
3. For each preset header, find the column index in `rows[h]` via normalized comparison: lowercase + trim + strip a single trailing `*` on both sides. If no match, emit a blank column with the preset's header text in row 0 (matches the existing `keepColumnsInOrder` "missing header" behavior).
4. Build the output: row 0 = preset header strings; rows 1..n = source rows `h+1..end` projected through the resolved column indices.
5. Clear the used range. Write the output to a range anchored at `A1` sized to the new dimensions.

Single Excel sync at the end. No sort step. No row filtering.

## Custom section behavior

The "Keep & Reorder", "Keep (preserve order)", and "Remove" buttons stay in the UI and use the same header-row detection so they also work on multi-header reports. `keepColumnsBySet` and `removeColumnsByHeader` previously deleted columns by index after reading row 0; they now read row `h` instead and delete columns whose row-`h` value matches the user's list. The grouping row (if any) is left in place by these two operations — only `keepColumnsInOrder` rewrites the entire range, and only it discards the grouping row.

This keeps Custom's behavior conservative: it edits in place, doesn't reshape the sheet, and preserves whatever non-data structure already exists.

## File layout

The current `taskpane.ts` mixes UI wiring, preset config, theme, and Excel operations into one ~200-line file. Splitting:

- `src/taskpane/presets.ts` — `Preset` type and the `PRESETS` array. Pure data, no Office API.
- `src/taskpane/excel-ops.ts` — `detectHeaderRow`, `resolveColumns`, `keepColumnsInOrder`, `keepColumnsBySet`, `removeColumnsByHeader`, plus a `setStatus` helper hook (passed in or imported from a small `status.ts`). All `Excel.run` calls live here.
- `src/taskpane/taskpane.ts` — `Office.onReady`, theme detection, button rendering from `PRESETS`, Custom section input parsing, status DOM.

Webpack entry stays at `taskpane.ts`; the others are imported.

## Out of scope (explicit non-goals)

- Star Achiever button and its previous-quarter logic.
- Teams button.
- Sort capability of any kind. The spec for the source data preserves the natural order each report ships in (Court of Sales is pre-sorted by Rank, Wholesale Scoreboard by Unit Wholesale descending, etc.).
- Row filtering (Total rows, blank-cell rows). No in-scope report has these in a way that affects the output.
- Custom section gaining preset-aware multi-row writing. Custom edits in place; it doesn't try to drop grouping rows.

## Risks and mitigations

- **Header drift.** Mary Kay InTouch could change a column name (e.g., drop the `*` from `Wholesale $ + Team Building*`, or expand `Conslt Name` to `Consultant Name`). Mitigation: the matcher's case-insensitive + trim + strip-trailing-`*` normalization absorbs the most common drift forms. If a header changes more substantively, the button emits a blank column with the configured header text and a status message will show how many columns were resolved vs. blank.
- **Missing header indication.** Today the code silently inserts a blank column when a header isn't found. The status line should mention the count of unresolved headers so users notice when a report's columns have shifted. Implementation will include this.
- **Used range includes trailing junk.** Some reports include trailing summary or blank rows. Out of scope to handle — `getUsedRange()` is what the user sees, and the existing tool already operates on it. If this becomes a real problem we'll address it specifically.
