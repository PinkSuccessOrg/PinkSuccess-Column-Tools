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
