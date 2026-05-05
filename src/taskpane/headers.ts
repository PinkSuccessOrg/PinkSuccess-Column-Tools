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

export function detectHeaderRow(rows: unknown[][]): HeaderLayout {
  if (rows.length < 2) return { headerRowIdx: 0, groupingRowIdx: -1 };
  const r0 = countNonEmptyCells(rows[0]);
  const r1 = countNonEmptyCells(rows[1]);
  if (r1 > r0) return { headerRowIdx: 1, groupingRowIdx: 0 };
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
