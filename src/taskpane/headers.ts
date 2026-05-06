export function normalizeHeader(value: unknown): string {
  if (value === null || value === undefined) return "";
  const s = String(value).trim().toLowerCase();
  return s.endsWith("*") ? s.slice(0, -1).trimEnd() : s;
}

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

// Row 1 is treated as the header row only when it has strictly more non-empty
// cells than row 0 by this margin. Without a margin, reports whose header row
// happens to have one missing/blank cell (e.g., InTouch's Court of Sharing
// where column 5 is unlabeled) get misclassified — the data row "wins" by 1
// non-empty cell and we can't resolve any preset headers.
const MULTIROW_HEADER_MARGIN = 2;

export function detectHeaderRow(rows: unknown[][]): HeaderLayout {
  if (rows.length < 2) return { headerRowIdx: 0, groupingRowIdx: -1 };
  const r0 = countNonEmptyCells(rows[0]);
  const r1 = countNonEmptyCells(rows[1]);
  if (r1 >= r0 + MULTIROW_HEADER_MARGIN) return { headerRowIdx: 1, groupingRowIdx: 0 };
  return { headerRowIdx: 0, groupingRowIdx: -1 };
}

export function resolveColumnIndex(headerRow: unknown[], wanted: string): number {
  const target = normalizeHeader(wanted);
  for (let i = 0; i < headerRow.length; i++) {
    if (normalizeHeader(headerRow[i]) === target) return i;
  }
  return -1;
}

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
