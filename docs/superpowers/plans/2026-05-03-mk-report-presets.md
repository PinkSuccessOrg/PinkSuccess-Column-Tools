# MK Report Presets Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the three placeholder presets in the PinkSuccess Column Tools task pane with eight Mary Kay report buttons sourced from `NL Report Fields.xlsx`, with multi-row-header detection so reports like Star Consultant Tracking and Team Commissions work correctly.

**Architecture:** Split the existing 200-line `taskpane.ts` into four focused files: pure header utilities (`headers.ts`), preset config (`presets.ts`), Excel.run wrappers (`excel-ops.ts`), and UI bootstrap (`taskpane.ts`). Add a minimal `node:test`-based test runner via `tsx` so the pure utilities can be unit-tested without an Excel host.

**Tech Stack:** TypeScript, Office.js Excel API, Webpack, Babel, `tsx` (new dev dep), Node.js built-in `node:test` + `node:assert`.

**Spec:** `docs/superpowers/specs/2026-05-03-mk-report-presets-design.md`

---

## File Structure

| Path | Status | Responsibility |
|---|---|---|
| `src/taskpane/headers.ts` | New | Pure header utilities: `normalizeHeader`, `detectHeaderRow`, `resolveColumnIndex`. No Office dependency. |
| `src/taskpane/headers.test.ts` | New | `node:test` unit tests for the pure utilities. |
| `src/taskpane/presets.ts` | New | `Preset` type and the eight `PRESETS`. No Office dependency. |
| `src/taskpane/excel-ops.ts` | New | `Excel.run` wrappers: `runPreset`, `keepColumnsInOrder`, `keepColumnsBySet`, `removeColumnsByHeader`. Uses `headers.ts` + `presets.ts`. |
| `src/taskpane/taskpane.ts` | Modify | UI bootstrap only: `Office.onReady`, theme, button rendering, custom-section input parsing, status DOM. Imports from the three new modules. |
| `src/taskpane/taskpane.html` | Unchanged | Layout already supports a flex-wrap `#presets` container and the Custom section. |
| `package.json` | Modify | Add `tsx` dev dep and `test` script. |
| `tsconfig.json` | Unchanged | Existing config compiles `.ts` test files via `tsx`. |

---

## Task 1: Add test infrastructure

**Why:** The pure utilities in `headers.ts` need fast feedback while we build them. Office add-ins can't be unit-tested without a heavy mock layer, but our pure functions don't need the Office API at all. Adding `tsx` lets `node:test` run TypeScript files directly with no compile step or Jest config.

**Files:**
- Modify: `package.json`
- Create: `src/taskpane/smoke.test.ts`

- [ ] **Step 1.1: Install `tsx` as a dev dependency**

Run from the project root:
```bash
npm install --save-dev tsx
```
Expected: `tsx` appears under `devDependencies` in `package.json`. No other package changes.

- [ ] **Step 1.2: Add a `test` script to `package.json`**

In `package.json` under `"scripts"`, add this entry (alphabetical order is fine, after `start` and before `stop`):

```json
"test": "node --import tsx --test 'src/taskpane/**/*.test.ts'",
```

The full `scripts` block should now include the existing entries plus this new one.

- [ ] **Step 1.3: Write a smoke test to verify the runner works**

Create `src/taskpane/smoke.test.ts` with this exact content:

```ts
import { test } from "node:test";
import assert from "node:assert/strict";

test("test runner is wired up", () => {
  assert.equal(1 + 1, 2);
});
```

- [ ] **Step 1.4: Run the smoke test**

Run:
```bash
npm test
```
Expected output includes `# pass 1` and exit code 0. If the runner errors with `Unknown file extension ".ts"`, `tsx` is not loading — re-check the `--import tsx` flag in the script.

- [ ] **Step 1.5: Delete the smoke test**

```bash
rm "src/taskpane/smoke.test.ts"
```

- [ ] **Step 1.6: Commit**

```bash
git add package.json package-lock.json
git commit -m "chore: add tsx + node:test runner for pure-function unit tests"
```

---

## Task 2: Pure header utilities (TDD)

**Why:** `normalizeHeader`, `detectHeaderRow`, and `resolveColumnIndex` encode subtle rules — case-insensitive matching, trailing-`*` tolerance, picking the row with more non-empty cells. Subtle rules without tests rot silently. These functions take plain arrays as input and return plain values, so testing them directly is straightforward.

**Files:**
- Create: `src/taskpane/headers.ts`
- Create: `src/taskpane/headers.test.ts`

- [ ] **Step 2.1: Write the failing test for `normalizeHeader`**

Create `src/taskpane/headers.test.ts` with this content:

```ts
import { test } from "node:test";
import assert from "node:assert/strict";
import { normalizeHeader, detectHeaderRow, resolveColumnIndex } from "./headers";

test("normalizeHeader: lowercases and trims", () => {
  assert.equal(normalizeHeader("  Name  "), "name");
});

test("normalizeHeader: strips a single trailing asterisk", () => {
  assert.equal(normalizeHeader("Wholesale $ + Team Building*"), "wholesale $ + team building");
});

test("normalizeHeader: leaves internal asterisks alone", () => {
  assert.equal(normalizeHeader("a*b"), "a*b");
});

test("normalizeHeader: handles non-string input", () => {
  assert.equal(normalizeHeader(null), "");
  assert.equal(normalizeHeader(undefined), "");
  assert.equal(normalizeHeader(42), "42");
});
```

- [ ] **Step 2.2: Run the test and verify it fails**

Run:
```bash
npm test
```
Expected: failures with messages mentioning `Cannot find module './headers'` or similar import errors.

- [ ] **Step 2.3: Implement `normalizeHeader`**

Create `src/taskpane/headers.ts` with this content:

```ts
export function normalizeHeader(value: unknown): string {
  if (value === null || value === undefined) return "";
  const s = String(value).trim().toLowerCase();
  return s.endsWith("*") ? s.slice(0, -1).trimEnd() : s;
}
```

- [ ] **Step 2.4: Run tests and verify `normalizeHeader` passes**

```bash
npm test
```
Expected: 4 passing, but `detectHeaderRow` and `resolveColumnIndex` tests still fail (because they aren't written or implemented yet — actually they aren't *imported* from a non-existent symbol, so this run will also error on the import line). To narrow scope, comment out the import names you haven't implemented or use `--test-name-pattern`. Simplest: temporarily edit the import line to just `import { normalizeHeader } from "./headers";` and remove unused names. Re-add them as you go.

Re-run after trimming the import:
```bash
npm test
```
Expected: 4 passing.

- [ ] **Step 2.5: Add failing tests for `detectHeaderRow`**

Append these tests to `src/taskpane/headers.test.ts` (and update the import to add `detectHeaderRow`):

```ts
test("detectHeaderRow: single header row in row 0", () => {
  const rows = [
    ["Name", "Birth Date"],
    ["Jane Doe", "1990-01-01"],
  ];
  assert.deepEqual(detectHeaderRow(rows), { headerRowIdx: 0, groupingRowIdx: -1 });
});

test("detectHeaderRow: grouping row 0 is mostly whitespace, headers in row 1", () => {
  const rows = [
    [" ", " ", " ", " ", "Personal Team", "Personal Team", "Personal Team"],
    ["Consultant Number", "First Name", "Last Name", "Name", "Wholesale", "%", "Team Commission"],
    ["6185EE", "Eileen", "Eckhoff", "Eileen Eckhoff", 264, 0.04, 10.56],
  ];
  assert.deepEqual(detectHeaderRow(rows), { headerRowIdx: 1, groupingRowIdx: 0 });
});

test("detectHeaderRow: empty input", () => {
  assert.deepEqual(detectHeaderRow([]), { headerRowIdx: 0, groupingRowIdx: -1 });
});

test("detectHeaderRow: only one row", () => {
  assert.deepEqual(detectHeaderRow([["A", "B", "C"]]), { headerRowIdx: 0, groupingRowIdx: -1 });
});

test("detectHeaderRow: ties go to row 0", () => {
  const rows = [
    ["A", "B", "C"],
    ["x", "y", "z"],
  ];
  assert.deepEqual(detectHeaderRow(rows), { headerRowIdx: 0, groupingRowIdx: -1 });
});
```

- [ ] **Step 2.6: Run and verify the new tests fail**

```bash
npm test
```
Expected: import error or `detectHeaderRow is not a function`.

- [ ] **Step 2.7: Implement `detectHeaderRow`**

Append to `src/taskpane/headers.ts`:

```ts
export interface HeaderLayout {
  headerRowIdx: number;
  groupingRowIdx: number;
}

function countNonEmptyCells(row: unknown[]): number {
  let n = 0;
  for (const cell of row) {
    if (cell === null || cell === undefined) continue;
    if (String(cell).trim() === "") continue;
    n++;
  }
  return n;
}

export function detectHeaderRow(rows: unknown[][]): HeaderLayout {
  if (rows.length < 2) return { headerRowIdx: 0, groupingRowIdx: -1 };
  const r0 = countNonEmptyCells(rows[0]);
  const r1 = countNonEmptyCells(rows[1]);
  if (r1 > r0) return { headerRowIdx: 1, groupingRowIdx: 0 };
  return { headerRowIdx: 0, groupingRowIdx: -1 };
}
```

- [ ] **Step 2.8: Run tests and verify all pass**

```bash
npm test
```
Expected: 9 passing (4 normalize + 5 detect).

- [ ] **Step 2.9: Add failing tests for `resolveColumnIndex`**

Append to `src/taskpane/headers.test.ts` (and update the import to add `resolveColumnIndex`):

```ts
test("resolveColumnIndex: exact case-sensitive match", () => {
  const headerRow = ["Name", "Birth Date"];
  assert.equal(resolveColumnIndex(headerRow, "Name"), 0);
  assert.equal(resolveColumnIndex(headerRow, "Birth Date"), 1);
});

test("resolveColumnIndex: case-insensitive match", () => {
  const headerRow = ["Number Of Years"];
  assert.equal(resolveColumnIndex(headerRow, "Number of Years"), 0);
});

test("resolveColumnIndex: tolerates trailing asterisk on source", () => {
  const headerRow = ["Wholesale $ + Team Building*"];
  assert.equal(resolveColumnIndex(headerRow, "Wholesale $ + Team Building"), 0);
});

test("resolveColumnIndex: tolerates trailing asterisk on preset", () => {
  const headerRow = ["Wholesale $ + Team Building"];
  assert.equal(resolveColumnIndex(headerRow, "Wholesale $ + Team Building*"), 0);
});

test("resolveColumnIndex: returns -1 when not found", () => {
  const headerRow = ["Name", "Birth Date"];
  assert.equal(resolveColumnIndex(headerRow, "Phone"), -1);
});

test("resolveColumnIndex: matches the percent-sign header", () => {
  const headerRow = ["Name", "%", "Team Commission"];
  assert.equal(resolveColumnIndex(headerRow, "%"), 1);
});
```

- [ ] **Step 2.10: Run and verify the new tests fail**

```bash
npm test
```
Expected: import error or `resolveColumnIndex is not a function`.

- [ ] **Step 2.11: Implement `resolveColumnIndex`**

Append to `src/taskpane/headers.ts`:

```ts
export function resolveColumnIndex(headerRow: unknown[], wanted: string): number {
  const target = normalizeHeader(wanted);
  for (let i = 0; i < headerRow.length; i++) {
    if (normalizeHeader(headerRow[i]) === target) return i;
  }
  return -1;
}
```

- [ ] **Step 2.12: Run tests and verify all pass**

```bash
npm test
```
Expected: 15 passing.

- [ ] **Step 2.13: Commit**

```bash
git add src/taskpane/headers.ts src/taskpane/headers.test.ts
git commit -m "feat: add pure header utilities with multi-row + asterisk-tolerant matching"
```

---

## Task 3: Preset config module

**Why:** Pulling the eight presets out of `taskpane.ts` into `presets.ts` keeps the data declarative and easy to edit. No Office dependency means it can be imported anywhere safely.

**Files:**
- Create: `src/taskpane/presets.ts`

- [ ] **Step 3.1: Create `src/taskpane/presets.ts`**

Write this exact content:

```ts
export interface Preset {
  id: string;
  label: string;
  headers: string[];
}

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
  },
  {
    id: "preset-ct-sales",
    label: "Ct Sales",
    headers: ["Rank", "Conslt Name", "YTD Total"],
  },
  {
    id: "preset-ct-sharing",
    label: "Ct Sharing",
    headers: ["Name", "Sem Qual Team Member", "Sem Recruiter Comm Earned"],
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
  },
  {
    id: "preset-whsl",
    label: "WHSL",
    headers: ["Name", "Unit Wholesale"],
  },
  {
    id: "preset-checks",
    label: "Checks",
    headers: ["Name", "%", "Team Commission"],
  },
  {
    id: "preset-new-cons",
    label: "New Cons",
    headers: ["New Consultant", "From", "Recruiter"],
  },
];
```

- [ ] **Step 3.2: Verify TypeScript compiles**

Run:
```bash
npx tsc --noEmit
```
Expected: no errors. (`noEmit` validates types without writing files.)

- [ ] **Step 3.3: Commit**

```bash
git add src/taskpane/presets.ts
git commit -m "feat: extract preset config into presets.ts with 8 MK report buttons"
```

---

## Task 4: Excel ops module

**Why:** Centralizing all `Excel.run` calls in one file keeps the side-effecting code separate from pure logic and from UI wiring. Each operation reads the used range, detects header layout, transforms, and writes back.

**Files:**
- Create: `src/taskpane/excel-ops.ts`

- [ ] **Step 4.1: Create `src/taskpane/excel-ops.ts`**

Write this exact content:

```ts
/* global Excel */

import { detectHeaderRow, resolveColumnIndex } from "./headers";
import { Preset } from "./presets";

export type StatusFn = (msg: string) => void;

export async function runPreset(preset: Preset, setStatus: StatusFn): Promise<void> {
  await keepColumnsInOrder(preset.headers, setStatus);
}

export async function keepColumnsInOrder(
  keepHeaders: string[],
  setStatus: StatusFn
): Promise<void> {
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

      const outputRows: unknown[][] = [];
      outputRows.push(keepHeaders.slice());
      for (let r = dataStart; r < data.length; r++) {
        const srcRow = data[r];
        outputRows.push(sourceCols.map((idx) => (idx === -1 ? "" : srcRow[idx])));
      }

      used.clear();
      const target = sheet.getRangeByIndexes(0, 0, outputRows.length, keepHeaders.length);
      target.values = outputRows as Excel.Range["values"];
      await context.sync();

      const missing = sourceCols.filter((i) => i === -1).length;
      const summary = `Kept ${keepHeaders.length - missing}/${keepHeaders.length} column(s)` +
        (missing > 0 ? ` (${missing} not found in source — emitted blank)` : "");
      setStatus(summary + ".");
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
  const wanted = new Set(keepHeaders.map((h) => h.toLowerCase()));
  await deleteWhere((header) => !wanted.has(header.toLowerCase()), "Kept", setStatus);
}

export async function removeColumnsByHeader(
  removeHeaders: string[],
  setStatus: StatusFn
): Promise<void> {
  if (removeHeaders.length === 0) {
    setStatus("No headers given.");
    return;
  }
  const drop = new Set(removeHeaders.map((h) => h.toLowerCase()));
  await deleteWhere((header) => drop.has(header.toLowerCase()), "Removed", setStatus);
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

- [ ] **Step 4.2: Verify TypeScript compiles**

```bash
npx tsc --noEmit
```
Expected: no errors.

- [ ] **Step 4.3: Commit**

```bash
git add src/taskpane/excel-ops.ts
git commit -m "feat: add excel-ops module with header-layout-aware operations"
```

---

## Task 5: Refactor `taskpane.ts` to use the new modules

**Why:** With the data and Excel logic extracted, `taskpane.ts` shrinks to UI-only concerns: Office bootstrap, theme, button rendering, custom-input parsing, status DOM.

**Files:**
- Modify: `src/taskpane/taskpane.ts` (full rewrite)

- [ ] **Step 5.1: Replace `src/taskpane/taskpane.ts` with the slim version**

Overwrite the file with this exact content:

```ts
/* global Office */

import { PRESETS } from "./presets";
import {
  runPreset,
  keepColumnsInOrder,
  keepColumnsBySet,
  removeColumnsByHeader,
} from "./excel-ops";

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  applyTheme();

  const container = document.getElementById("presets")!;
  for (const p of PRESETS) {
    const btn = document.createElement("button");
    btn.className = "ms-Button";
    btn.id = p.id;
    btn.textContent = p.label;
    btn.addEventListener("click", () => void runPreset(p, setStatus));
    container.appendChild(btn);
  }

  bindClick("run-keep-order", () => keepColumnsInOrder(parseCustom(), setStatus));
  bindClick("run-keep-set", () => keepColumnsBySet(parseCustom(), setStatus));
  bindClick("run-remove", () => removeColumnsByHeader(parseCustom(), setStatus));
});

// ---------- Theme ----------

function applyTheme(): void {
  const theme = Office.context.officeTheme;
  const dark =
    theme && theme.bodyBackgroundColor
      ? isDarkColor(theme.bodyBackgroundColor)
      : window.matchMedia("(prefers-color-scheme: dark)").matches;
  document.body.classList.toggle("theme-dark", dark);
  document.body.classList.toggle("theme-light", !dark);
}

function isDarkColor(hex: string): boolean {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) return false;
  const v = parseInt(m[1], 16);
  const r = (v >> 16) & 255;
  const g = (v >> 8) & 255;
  const b = v & 255;
  return 0.299 * r + 0.587 * g + 0.114 * b < 128;
}

// ---------- DOM helpers ----------

function bindClick(id: string, handler: () => Promise<void> | void): void {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element: ${id}`);
  el.addEventListener("click", () => {
    void handler();
  });
}

function parseCustom(): string[] {
  const raw = (document.getElementById("custom-headers") as HTMLInputElement).value;
  return raw
    .split(",")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

function setStatus(msg: string): void {
  const el = document.getElementById("status");
  if (el) el.textContent = msg;
}
```

- [ ] **Step 5.2: Verify TypeScript compiles**

```bash
npx tsc --noEmit
```
Expected: no errors.

- [ ] **Step 5.3: Verify a production build succeeds**

```bash
npm run build
```
Expected: webpack output ending in something like `webpack 5.x compiled successfully`. No TypeScript or module-resolution errors. Webpack writes to `dist/` (existing config).

- [ ] **Step 5.4: Verify the unit tests still pass**

```bash
npm test
```
Expected: 15 passing.

- [ ] **Step 5.5: Commit**

```bash
git add src/taskpane/taskpane.ts
git commit -m "refactor: slim taskpane.ts down to UI bootstrap; presets driven from PRESETS array"
```

---

## Task 6: Manual verification in Excel

**Why:** The unit tests cover header utilities, but the Excel.run path can only be verified against a real workbook. Validate every preset against the matching example file from `~/Downloads/mk_report_examples/`.

This task has no automated test step. Each substep is a manual check; mark it done after you observe the expected behavior.

**Files:** none modified.

- [ ] **Step 6.1: Start the dev server**

```bash
npm run dev-server
```
Leave it running. It serves `https://localhost:3000/taskpane.html`.

- [ ] **Step 6.2: Sideload the manifest in Excel**

In a separate terminal:
```bash
npm start
```
Excel should launch with the add-in pinned. (If not, follow the prompts the Office tooling prints.)

- [ ] **Step 6.3: Verify all 8 buttons render**

Open the task pane. Under "Presets", expect to see 8 buttons in this order: `OT Star`, `Ct Sales`, `Ct Sharing`, `B-Day`, `Anniv`, `WHSL`, `Checks`, `New Cons`. The old `Customer Export` / `Sales Report` / `Strip Internal Fields` buttons should be gone.

- [ ] **Step 6.4: Verify OT Star against `Star Consultant Tracking.xlsx`**

Open `~/Downloads/mk_report_examples/Star Consultant Tracking.xlsx`. This file has a multi-row header (row 1 = `Contest Credit Needed` grouping, row 2 = real headers). Click `OT Star`. Expect:
- The sheet now has exactly 7 columns: `Name`, `Wholesale $ + Team Building`, `Sapphire`, `Ruby`, `Diamond`, `Emerald`, `Pearl` (in that order).
- Row 1 contains those headers.
- Rows 2+ contain the data rows previously in rows 3+ of the source.
- The status line reads `Kept 7/7 column(s).`
- The grouping row is gone.

If any column is blank or the grouping row remains, stop and inspect.

- [ ] **Step 6.5: Verify Ct Sales against `Unit Recognition Seminar YTD Court of Sales.xlsx`**

Open the file. Click `Ct Sales`. Expect 3 columns: `Rank`, `Conslt Name`, `YTD Total`. Status: `Kept 3/3 column(s).` Note the preset spec uses the actual source header `Conslt Name` (not "Consultant Name"). Data should remain pre-sorted by `Rank`.

- [ ] **Step 6.6: Verify Ct Sharing against `Unit Recognition Seminar YTD Court of Sharing.xlsx`**

Open the file. Click `Ct Sharing`. Expect 3 columns: `Name`, `Sem Qual Team Member`, `Sem Recruiter Comm Earned`. Status: `Kept 3/3 column(s).`

- [ ] **Step 6.7: Verify B-Day against `Birthdays, Anniversaries and Addresses-Unit.xlsx` (the file *without* `(1)`)**

Open the file. Click `B-Day`. Expect 2 columns: `Name`, `Birth Date`. Status: `Kept 2/2 column(s).`

- [ ] **Step 6.8: Verify Anniv against `Birthdays, Anniversaries and Addresses-Unit (1).xlsx`**

Open the file. Click `Anniv`. Expect 2 columns: `Name`, `Number Of Years`. Status: `Kept 2/2 column(s).` (The source has `Number Of Years` with capital `O`; case-insensitive match handles this.)

- [ ] **Step 6.9: Verify WHSL against `Unit Recognition Unit Wholesale Scoreboard.xlsx`**

Open the file. Click `WHSL`. Expect 2 columns: `Name`, `Unit Wholesale`. Status: `Kept 2/2 column(s).`

- [ ] **Step 6.10: Verify Checks against `Team Commissions-Unit.xlsx`**

Open the file. This has a multi-row header (row 1 = `Personal Team` grouping, row 2 = real headers). Click `Checks`. Expect 3 columns: `Name`, `%`, `Team Commission`. Status: `Kept 3/3 column(s).` Grouping row removed.

- [ ] **Step 6.11: Verify New Cons against `Unit Recognition New Unit Members.xlsx`**

Open the file. Click `New Cons`. Expect 3 columns: `New Consultant`, `From`, `Recruiter`. Status: `Kept 3/3 column(s).` (The example file has only the header row; data rows may be empty — that's fine, the headers should still be present.)

- [ ] **Step 6.12: Verify the Custom section still works**

In any open report, type `Name, Birth Date` (or any two real headers) into the custom-headers input. Click `Keep & Reorder`. Expect those two columns to remain. Then undo (Ctrl/Cmd+Z) and try `Remove` with one header — expect that column to be gone.

- [ ] **Step 6.13: Stop the dev server**

```bash
npm stop
```
And Ctrl-C the `npm run dev-server` terminal.

- [ ] **Step 6.14: Commit any verification notes (only if changes were needed)**

If verification surfaced bugs that were fixed in this session, commit them as separate fix commits. If nothing changed, no commit is needed for this task.

---

## Self-Review Checklist (already run by author)

- **Spec coverage:**
  - Eight presets with correct headers — Task 3.
  - Multi-row header detection — Task 2.
  - Asterisk-tolerant matching — Task 2.
  - File split into `presets.ts` / `excel-ops.ts` / `taskpane.ts` (+ `headers.ts` for testability) — Tasks 2-5.
  - Custom section uses the same header detection — Task 4 `keepColumnsBySet` / `removeColumnsByHeader`.
  - Status message reports unresolved columns — Task 4 `keepColumnsInOrder` summary.
  - Star Achiever and Teams omitted — confirmed by Task 3 PRESETS list.
- **Placeholder scan:** None. All code blocks are complete.
- **Type consistency:** `Preset` (id, label, headers) is referenced identically in `presets.ts`, `excel-ops.ts`, and `taskpane.ts`. `StatusFn` is consistent. `HeaderLayout` returns `headerRowIdx` and `groupingRowIdx` matching test expectations.
