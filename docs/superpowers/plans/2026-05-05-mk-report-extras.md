# MK Report Extras Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** After each preset extracts the requested columns, also apply universal styling (Avenir Next LT Pro 12, no borders, drop the header row) and per-preset column formatting (currency / center / left alignment).

**Architecture:** A new `ColumnFormatRule[]` field on `Preset` declares per-column rules. `keepColumnsInOrder` in `excel-ops.ts` is refactored to accept an options object, drop the header row from the written output, and run a formatting pass after the values are written. Universal rules (font + no borders) are hard-coded constants in `excel-ops.ts`. A small pure helper `expandColumnRule` lives in `headers.ts` so the index-resolution logic stays unit-testable.

**Tech Stack:** TypeScript, Office.js Excel API, `node:test` via `tsx`.

**Spec:** `docs/superpowers/specs/2026-05-05-mk-report-extras-design.md`

---

## File Structure

| Path | Status | Responsibility |
|---|---|---|
| `src/taskpane/headers.ts` | Modify | Add `expandColumnRule(columns, columnCount)` pure helper. |
| `src/taskpane/headers.test.ts` | Modify | Add unit tests for `expandColumnRule`. |
| `src/taskpane/presets.ts` | Modify | Add `Alignment`, `ColumnFormatRule` types and `columnFormats` field on `Preset`. Populate rules for OT Star, Ct Sales, Ct Sharing, Anniv, WHSL, Checks. |
| `src/taskpane/excel-ops.ts` | Modify | Refactor `keepColumnsInOrder` to take options object; drop header row from output; apply universal font + no borders + per-rule formatting. New helpers `applyUniversalFormat`, `applyColumnFormats`, `clearAllBorders`. |
| `src/taskpane/taskpane.ts` | Modify | Update Custom-section call site to pass the new options shape. |
| `public/index.html` | Modify | One-line note that Windows users may need to install Avenir Next LT Pro. |

---

## Task 1: Pure helper — `expandColumnRule`

**Why:** Translating a `ColumnFormatRule['columns']` (a single index, or an inclusive `[start, end]` tuple) into an array of clamped indices is pure logic — easy to test in isolation. Keeping it out of `excel-ops.ts` lets us catch off-by-one and clamping bugs before the Excel-API layer ever runs.

**Files:**
- Modify: `src/taskpane/headers.ts`
- Modify: `src/taskpane/headers.test.ts`

- [ ] **Step 1.1: Write the failing tests**

Append to `src/taskpane/headers.test.ts`. Update the import line at the top to include `expandColumnRule`:

```ts
import { normalizeHeader, detectHeaderRow, resolveColumnIndex, expandColumnRule } from "./headers";
```

Then append:

```ts
test("expandColumnRule: single index returns one-element array", () => {
  assert.deepEqual(expandColumnRule(2, 5), [2]);
});

test("expandColumnRule: inclusive range expands all indices", () => {
  assert.deepEqual(expandColumnRule([1, 4], 7), [1, 2, 3, 4]);
});

test("expandColumnRule: clamps the end of a range to columnCount-1", () => {
  assert.deepEqual(expandColumnRule([1, 10], 5), [1, 2, 3, 4]);
});

test("expandColumnRule: clamps a single index past the end to empty", () => {
  assert.deepEqual(expandColumnRule(7, 3), []);
});

test("expandColumnRule: negative single index is empty", () => {
  assert.deepEqual(expandColumnRule(-1, 5), []);
});

test("expandColumnRule: range with start > columnCount is empty", () => {
  assert.deepEqual(expandColumnRule([10, 20], 5), []);
});

test("expandColumnRule: range with start > end is empty", () => {
  assert.deepEqual(expandColumnRule([4, 1], 7), []);
});
```

- [ ] **Step 1.2: Run tests and verify they fail**

Run from project root:
```
npm test
```
Expected: failures with `expandColumnRule is not a function` (or similar import error).

- [ ] **Step 1.3: Implement `expandColumnRule`**

Append to `src/taskpane/headers.ts`:

```ts
export function expandColumnRule(
  columns: number | [number, number],
  columnCount: number
): number[] {
  if (typeof columns === "number") {
    return columns >= 0 && columns < columnCount ? [columns] : [];
  }
  const [rawStart, rawEnd] = columns;
  const start = Math.max(0, rawStart);
  const end = Math.min(columnCount - 1, rawEnd);
  if (start > end) return [];
  const out: number[] = [];
  for (let i = start; i <= end; i++) out.push(i);
  return out;
}
```

- [ ] **Step 1.4: Run tests and verify they pass**

Run:
```
npm test
```
Expected: 22 passing (15 existing + 7 new).

- [ ] **Step 1.5: Commit**

```
git add src/taskpane/headers.ts src/taskpane/headers.test.ts
git commit -m "feat(headers): add expandColumnRule pure helper for column-format index resolution"
```

---

## Task 2: Preset config — types + per-preset column formats

**Why:** Adds the declarative data the formatting pass will consume. Pure data, no Office API.

**Files:**
- Modify: `src/taskpane/presets.ts`

- [ ] **Step 2.1: Replace `src/taskpane/presets.ts` with the new content**

Overwrite the file with this exact content:

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

const CURRENCY = "$#,##0.00";

export const PRESETS: Preset[] = [
  {
    id: "preset-ot-star",
    label: "OT Star",
    headers: [
      "Name",
      "Wholesale $ + Team Building",
      "Sapphire",
      "Ruby",
      "Diamond",
      "Emerald",
      "Pearl",
    ],
    columnFormats: [
      { columns: [1, 6], numberFormat: CURRENCY, horizontalAlignment: "Center" },
    ],
  },
  {
    id: "preset-ct-sales",
    label: "Ct Sales",
    headers: ["Rank", "Conslt Name", "YTD Total"],
    columnFormats: [
      { columns: 0, horizontalAlignment: "Left" },
      { columns: 2, numberFormat: CURRENCY },
    ],
  },
  {
    id: "preset-ct-sharing",
    label: "Ct Sharing",
    headers: ["Name", "Sem Qual Team Member", "Sem Recruiter Comm Earned"],
    columnFormats: [
      { columns: 2, numberFormat: CURRENCY },
    ],
  },
  {
    id: "preset-bday",
    label: "B-Day",
    headers: ["Name", "Birth Date"],
  },
  {
    id: "preset-anniv",
    label: "Anniv",
    headers: ["Name", "Number Of Years"],
    columnFormats: [
      { columns: 1, horizontalAlignment: "Left" },
    ],
  },
  {
    id: "preset-whsl",
    label: "WHSL",
    headers: ["Name", "Unit Wholesale"],
    columnFormats: [
      { columns: 1, numberFormat: CURRENCY },
    ],
  },
  {
    id: "preset-checks",
    label: "Checks",
    headers: ["Name", "%", "Team Commission"],
    columnFormats: [
      { columns: 1, horizontalAlignment: "Center" },
      { columns: 2, numberFormat: CURRENCY },
    ],
  },
  {
    id: "preset-new-cons",
    label: "New Cons",
    headers: ["New Consultant", "From", "Recruiter"],
  },
];
```

- [ ] **Step 2.2: Verify TypeScript compiles**

Run:
```
npx tsc --noEmit
```
Expected: zero output. `excel-ops.ts` will already pull `Preset` and now sees the new optional field — no error because it's not used yet.

- [ ] **Step 2.3: Verify tests still pass**

Run:
```
npm test
```
Expected: 22 passing.

- [ ] **Step 2.4: Commit**

```
git add src/taskpane/presets.ts
git commit -m "feat(presets): add ColumnFormatRule type and per-preset extras config"
```

---

## Task 3: `keepColumnsInOrder` — drop header row, accept options object

**Why:** Two changes that are best done together to avoid a broken intermediate state: (1) the function now takes `{ headers, columnFormats? }` instead of a bare `string[]`, so it can carry per-preset formatting through; (2) the output range no longer includes the header row, per the "no headers" universal rule.

This task does *not* yet apply the formatting — that's Task 4. The point of splitting is to keep each commit small and verifiable.

**Files:**
- Modify: `src/taskpane/excel-ops.ts`
- Modify: `src/taskpane/taskpane.ts`

- [ ] **Step 3.1: Update `excel-ops.ts` — change signature and drop the header row**

Replace the entire contents of `src/taskpane/excel-ops.ts` with:

```ts
/* global Excel */

import { detectHeaderRow, normalizeHeader, resolveColumnIndex } from "./headers";
import { ColumnFormatRule, Preset } from "./presets";

export type StatusFn = (msg: string) => void;

export interface KeepInOrderOptions {
  headers: string[];
  columnFormats?: ColumnFormatRule[];
}

export async function runPreset(preset: Preset, setStatus: StatusFn): Promise<void> {
  await keepColumnsInOrder(
    { headers: preset.headers, columnFormats: preset.columnFormats },
    setStatus
  );
}

export async function keepColumnsInOrder(
  opts: KeepInOrderOptions,
  setStatus: StatusFn
): Promise<void> {
  const { headers: keepHeaders } = opts;
  if (keepHeaders.length === 0) {
    setStatus("No headers given.");
    return;
  }
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();
      used.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const data = used.values as unknown[][];
      const layout = detectHeaderRow(data);
      const headerRow = (data[layout.headerRowIdx] ?? []).map((h) => String(h ?? ""));
      const dataStart = layout.headerRowIdx + 1;

      const sourceCols: number[] = keepHeaders.map((wanted) =>
        resolveColumnIndex(headerRow, wanted)
      );

      // Data-only output: the universal "no headers" rule means we never write the header row.
      const outputRows: unknown[][] = [];
      for (let r = dataStart; r < data.length; r++) {
        const srcRow = data[r];
        outputRows.push(sourceCols.map((idx) => (idx === -1 ? "" : srcRow[idx])));
      }

      used.clear();

      const missing = sourceCols.filter((i) => i === -1).length;
      const missingSuffix = missing > 0 ? ` (${missing} not found in source — emitted blank)` : "";

      if (outputRows.length === 0) {
        setStatus(`Kept ${keepHeaders.length - missing}/${keepHeaders.length} column(s)${missingSuffix} (no data rows).`);
        await context.sync();
        return;
      }

      const target = sheet.getRangeByIndexes(0, 0, outputRows.length, keepHeaders.length);
      target.values = outputRows as Excel.Range["values"];
      await context.sync();

      setStatus(`Kept ${keepHeaders.length - missing}/${keepHeaders.length} column(s)${missingSuffix}.`);
    });
  } catch (e) {
    setStatus(`Error: ${(e as Error).message}`);
  }
}

export async function keepColumnsBySet(
  keepHeaders: string[],
  setStatus: StatusFn
): Promise<void> {
  if (keepHeaders.length === 0) {
    setStatus("No headers given.");
    return;
  }
  const wanted = new Set(keepHeaders.map(normalizeHeader));
  await deleteWhere((header) => !wanted.has(normalizeHeader(header)), "Kept", setStatus);
}

export async function removeColumnsByHeader(
  removeHeaders: string[],
  setStatus: StatusFn
): Promise<void> {
  if (removeHeaders.length === 0) {
    setStatus("No headers given.");
    return;
  }
  const drop = new Set(removeHeaders.map(normalizeHeader));
  await deleteWhere((header) => drop.has(normalizeHeader(header)), "Removed", setStatus);
}

async function deleteWhere(
  shouldDelete: (header: string) => boolean,
  verb: string,
  setStatus: StatusFn
): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();
      used.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const data = used.values as unknown[][];
      const layout = detectHeaderRow(data);
      const headerRow = (data[layout.headerRowIdx] ?? []).map((h) => String(h ?? ""));
      let deleted = 0;

      // Iterate right-to-left so deleting column i doesn't invalidate indices < i.
      for (let i = headerRow.length - 1; i >= 0; i--) {
        if (shouldDelete(headerRow[i])) {
          sheet
            .getRangeByIndexes(0, i, used.rowCount, 1)
            .delete(Excel.DeleteShiftDirection.left);
          deleted++;
        }
      }
      await context.sync();
      setStatus(`${verb} ${deleted} column(s).`);
    });
  } catch (e) {
    setStatus(`Error: ${(e as Error).message}`);
  }
}
```

- [ ] **Step 3.2: Update `taskpane.ts` — Custom Keep & Reorder uses the new shape**

In `src/taskpane/taskpane.ts`, find this line:

```ts
  bindClick("run-keep-order", () => keepColumnsInOrder(parseCustom(), setStatus));
```

Replace it with:

```ts
  bindClick("run-keep-order", () => keepColumnsInOrder({ headers: parseCustom() }, setStatus));
```

The other two Custom buttons (`run-keep-set`, `run-remove`) keep their existing string-array signatures and stay unchanged.

- [ ] **Step 3.3: Verify TypeScript compiles**

Run:
```
npx tsc --noEmit
```
Expected: zero output.

- [ ] **Step 3.4: Verify tests still pass**

Run:
```
npm test
```
Expected: 22 passing.

- [ ] **Step 3.5: Verify production build succeeds**

Run:
```
npm run build
```
Expected: `webpack 5.x compiled successfully`.

- [ ] **Step 3.6: Commit**

```
git add src/taskpane/excel-ops.ts src/taskpane/taskpane.ts
git commit -m "feat(excel-ops): drop header row from output; keepColumnsInOrder takes options object"
```

---

## Task 4: Apply formatting — universal rules + per-column rules

**Why:** This is the actual "extras" feature. After the values are written, we apply font + no borders to the entire output range, then walk the per-preset rules to apply currency / alignment to specific column slices.

**Files:**
- Modify: `src/taskpane/excel-ops.ts`

- [ ] **Step 4.1: Add formatting constants and helpers**

In `src/taskpane/excel-ops.ts`, update the import line to also bring in `expandColumnRule`:

```ts
import { detectHeaderRow, expandColumnRule, normalizeHeader, resolveColumnIndex } from "./headers";
```

Then, just below the existing `import { ColumnFormatRule, Preset } from "./presets";` line, add:

```ts
const UNIVERSAL_FONT = "Avenir Next LT Pro";
const UNIVERSAL_FONT_SIZE = 12;

const BORDER_TYPES: Excel.BorderIndex[] = [
  "EdgeTop",
  "EdgeBottom",
  "EdgeLeft",
  "EdgeRight",
  "InsideHorizontal",
  "InsideVertical",
];

function applyUniversalFormat(range: Excel.Range): void {
  range.format.font.name = UNIVERSAL_FONT;
  range.format.font.size = UNIVERSAL_FONT_SIZE;
  for (const border of BORDER_TYPES) {
    range.format.borders.getItem(border).style = "None";
  }
}

function applyColumnFormats(
  sheet: Excel.Worksheet,
  rowCount: number,
  columnCount: number,
  rules: ColumnFormatRule[] | undefined
): void {
  if (!rules || rowCount === 0) return;
  for (const rule of rules) {
    const indices = expandColumnRule(rule.columns, columnCount);
    for (const idx of indices) {
      const colRange = sheet.getRangeByIndexes(0, idx, rowCount, 1);
      if (rule.numberFormat) {
        const formatGrid: string[][] = [];
        for (let r = 0; r < rowCount; r++) formatGrid.push([rule.numberFormat]);
        colRange.numberFormat = formatGrid;
      }
      if (rule.horizontalAlignment) {
        colRange.format.horizontalAlignment = rule.horizontalAlignment;
      }
    }
  }
}
```

- [ ] **Step 4.2: Wire the formatting pass into `keepColumnsInOrder`**

In the same file, find the block inside `keepColumnsInOrder`:

```ts
      const target = sheet.getRangeByIndexes(0, 0, outputRows.length, keepHeaders.length);
      target.values = outputRows as Excel.Range["values"];
      await context.sync();

      setStatus(`Kept ${keepHeaders.length - missing}/${keepHeaders.length} column(s)${missingSuffix}.`);
```

Replace it with:

```ts
      const target = sheet.getRangeByIndexes(0, 0, outputRows.length, keepHeaders.length);
      target.values = outputRows as Excel.Range["values"];

      applyUniversalFormat(target);
      applyColumnFormats(sheet, outputRows.length, keepHeaders.length, opts.columnFormats);

      await context.sync();

      setStatus(`Kept ${keepHeaders.length - missing}/${keepHeaders.length} column(s)${missingSuffix}.`);
```

- [ ] **Step 4.3: Verify TypeScript compiles**

Run:
```
npx tsc --noEmit
```
Expected: zero output. If `Excel.BorderIndex` causes a type error, change the type annotation to `string[]` — the type is `Excel.BorderIndex` in recent `@types/office-js` but older versions may not export it as a named type.

- [ ] **Step 4.4: Verify tests still pass**

Run:
```
npm test
```
Expected: 22 passing.

- [ ] **Step 4.5: Verify production build succeeds**

Run:
```
npm run build
```
Expected: `webpack 5.x compiled successfully`.

- [ ] **Step 4.6: Commit**

```
git add src/taskpane/excel-ops.ts
git commit -m "feat(excel-ops): apply universal font/no-borders + per-preset column formatting after preset run"
```

---

## Task 5: Documentation note about Avenir Next LT Pro on Windows

**Why:** Windows users don't have the font installed by default; their Excel will silently substitute. A small note on the landing page heads off the "the output looks weird" support question.

**Files:**
- Modify: `public/index.html`

- [ ] **Step 5.1: Add the font note to the landing page**

In `public/index.html`, find this line in the "Before you start" `<ul>`:

```html
      <li>Open the report you want to clean up in Excel <strong>before</strong> clicking a preset button. The add-in operates on whatever sheet is active.</li>
```

Append a new `<li>` immediately after it, inside the same `<ul>`:

```html
      <li>The add-in styles your output in <strong>Avenir Next LT Pro</strong> at 12pt. macOS has this font preinstalled. Windows does not &mdash; if your output is showing in the wrong font, install Avenir Next LT Pro from your usual font source (Adobe Fonts, Linotype, etc.) and restart Excel.</li>
```

- [ ] **Step 5.2: Verify the page builds correctly**

Run:
```
npm run build
```
Expected: `webpack 5.x compiled successfully`. `dist/index.html` should have grown by a few hundred bytes.

- [ ] **Step 5.3: Commit**

```
git add public/index.html
git commit -m "docs(pages): note that Windows users may need to install Avenir Next LT Pro"
```

---

## Task 6: Manual verification in Excel

**Why:** TypeScript compilation and unit tests cover the pure logic. The Excel-API formatting calls can only be verified against a real workbook. This task has no automation; mark each substep done after observing the expected behavior.

**Files:** none modified.

- [ ] **Step 6.1: Push the branch**

```
git push origin main
```

This kicks off the GitHub Pages deploy. Wait ~2 minutes, confirm the workflow ran successfully on the Actions tab.

- [ ] **Step 6.2: Reload the add-in in Excel**

The add-in should already be sideloaded from earlier work. The new build is fetched fresh each time the task pane opens (Excel cache aside). To force a refresh: close the task pane, then reopen via **Show Tools** on the Home tab.

If you don't see updated behavior, clear the Office cache (Mac: `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/`; Windows: see Microsoft Learn "Clear the Office cache"). Reopen Excel.

- [ ] **Step 6.3: Verify OT Star formatting**

Open `~/Downloads/mk_report_examples/Star Consultant Tracking.xlsx`. Click **OT Star**. Expect:
- Output is 7 columns (Name, Wholesale $ + Team Building, Sapphire, Ruby, Diamond, Emerald, Pearl).
- **No header row** — row 1 is the first data row (e.g., `Brigitte Iglay`, etc.).
- Cells render in Avenir Next LT Pro 12 (macOS) or a fallback font (Windows).
- No cell borders anywhere.
- Columns B–G show currency format (`$2,286.00`-style values) and are center-aligned.
- Status: `Kept 7/7 column(s).`

- [ ] **Step 6.4: Verify Ct Sales formatting**

Open `~/Downloads/mk_report_examples/Unit Recognition Seminar YTD Court of Sales.xlsx`. Click **Ct Sales**. Expect:
- 3 columns: Rank, Conslt Name, YTD Total. No header row.
- Column A (Rank) is left-aligned.
- Column C (YTD Total) shows currency (`$12,400.00`-style).
- Status: `Kept 3/3 column(s).`

- [ ] **Step 6.5: Verify Ct Sharing formatting**

Open `~/Downloads/mk_report_examples/Unit Recognition Seminar YTD Court of Sharing.xlsx`. Click **Ct Sharing**. Expect:
- 3 columns: Name, Sem Qual Team Member, Sem Recruiter Comm Earned. No header row.
- Column C shows currency.
- Status: `Kept 3/3 column(s).`

- [ ] **Step 6.6: Verify B-Day**

Open the Birthdays workbook (the one with a `Birth Date` column). Click **B-Day**. Expect:
- 2 columns: Name, Birth Date. No header row.
- No per-column formatting (B-Day has none); only universal font + no borders apply.
- Status: `Kept 2/2 column(s).`

- [ ] **Step 6.7: Verify Anniv**

Open the Anniversaries workbook (the `(1)` file with `Number Of Years`). Click **Anniv**. Expect:
- 2 columns: Name, Number Of Years. No header row.
- Column B left-aligned.
- Status: `Kept 2/2 column(s).`

- [ ] **Step 6.8: Verify WHSL**

Open `~/Downloads/mk_report_examples/Unit Recognition Unit Wholesale Scoreboard.xlsx`. Click **WHSL**. Expect:
- 2 columns: Name, Unit Wholesale. No header row.
- Column B currency.
- Status: `Kept 2/2 column(s).`

- [ ] **Step 6.9: Verify Checks**

Open `~/Downloads/mk_report_examples/Team Commissions-Unit.xlsx`. Click **Checks**. Expect:
- 3 columns: Name, %, Team Commission. No header row.
- Column B center-aligned.
- Column C currency.
- Status: `Kept 3/3 column(s).`

- [ ] **Step 6.10: Verify New Cons**

Open `~/Downloads/mk_report_examples/Unit Recognition New Unit Members.xlsx`. Click **New Cons**. Expect:
- 3 columns: New Consultant, From, Recruiter. No header row.
- No per-column formatting (New Cons has none); only universal font + no borders apply.
- Status: depends on whether the example file has data rows. The sample we saw earlier had only a header row → expect `Kept 3/3 column(s) (no data rows).`

- [ ] **Step 6.11: Verify Custom Keep & Reorder still works**

In any open report, type a real header into the custom-headers input (e.g., `Name, Birth Date`). Click **Keep & Reorder**. Expect:
- Output is just those columns.
- No header row (Custom Keep & Reorder gets the universal treatment per the spec).
- Universal font / no borders applied.

- [ ] **Step 6.12: Verify Custom Keep (preserve order) is unchanged**

In any open report, type a header into the input and click **Keep (preserve order)**. Expect:
- The original header row remains intact.
- Existing formatting (if any) on the kept columns is preserved.
- No font / border changes from this operation.

- [ ] **Step 6.13: Verify Custom Remove is unchanged**

Same as 6.12 but with **Remove**. Expect:
- The named columns are gone.
- The header row remains.
- Other formatting untouched.

- [ ] **Step 6.14: Commit any verification fixes if needed**

If verification surfaces a bug that gets fixed in this session, commit it as its own `fix:` commit referencing what broke. If everything passes, no commit is needed for this task.

---

## Self-Review Checklist (already run by author)

- **Spec coverage:**
  - Universal font + no borders + no headers — Tasks 3 (drop headers) and 4 (font, borders).
  - Per-preset column formatting — Tasks 2 (data) and 4 (apply).
  - `expandColumnRule` clamping — Task 1.
  - Currency `$#,##0.00` — Task 2 (`CURRENCY` constant).
  - Custom Keep & Reorder full treatment — Task 3 (signature change) + Task 4 (formatting pass applies because it goes through the same path).
  - Custom Keep (preserve order) and Remove unchanged — Task 3 (those callers untouched).
  - Avenir Next LT Pro Windows note — Task 5.
- **Placeholder scan:** None. Every code step shows complete code.
- **Type consistency:**
  - `Alignment` and `ColumnFormatRule` defined in Task 2, consumed in Task 4.
  - `KeepInOrderOptions` defined in Task 3, used by Custom call site in Task 3 and by `runPreset` in Task 3.
  - `expandColumnRule(columns, columnCount)` signature in Task 1 matches the call in Task 4.
  - `applyUniversalFormat(range)` and `applyColumnFormats(sheet, rowCount, columnCount, rules)` signatures defined in Task 4 Step 4.1 match the call sites in Task 4 Step 4.2.
