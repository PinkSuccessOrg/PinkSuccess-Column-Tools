/* global Excel */

import { detectHeaderRow, expandColumnRule, normalizeHeader, resolveColumnIndex } from "./headers";
import { ColumnFormatRule, Preset } from "./presets";

const UNIVERSAL_FONT = "Avenir Next LT Pro";
const UNIVERSAL_FONT_SIZE = 12;

const BORDER_TYPES: string[] = [
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
    range.format.borders.getItem(border as Excel.BorderIndex).style = "None";
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

export type StatusFn = (msg: string) => void;

export interface KeepInOrderOptions {
  headers: string[];
  // Optional sparse list of display labels parallel to `headers`. Falls back
  // to `headers[i]` when missing or undefined. Only used when keepHeader is true.
  headerLabels?: (string | undefined)[];
  // Default false: the universal "no headers" rule strips the header row.
  // Set true to write a header row above the data (e.g., OT Star).
  keepHeader?: boolean;
  columnFormats?: ColumnFormatRule[];
}

export async function runPreset(preset: Preset, setStatus: StatusFn): Promise<void> {
  await keepColumnsInOrder(
    {
      headers: preset.headers,
      headerLabels: preset.headerLabels,
      keepHeader: preset.keepHeader,
      columnFormats: preset.columnFormats,
    },
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
      used.load(["values", "text", "numberFormat", "valueTypes", "rowCount", "columnCount"]);
      await context.sync();

      const data = used.values as unknown[][];
      const sourceText = used.text as string[][];
      const sourceFormats = used.numberFormat as string[][];
      const sourceValueTypes = used.valueTypes as Excel.RangeValueType[][];
      const layout = detectHeaderRow(data);
      const headerRow = (data[layout.headerRowIdx] ?? []).map((h) => String(h ?? ""));
      const dataStart = layout.headerRowIdx + 1;

      const sourceCols: number[] = keepHeaders.map((wanted) =>
        resolveColumnIndex(headerRow, wanted)
      );

      // Target column indices flagged forceText — we'll substitute the source's
      // displayed text for the value and lock these cells as "@".
      const forceTextTargets = new Set<number>();
      for (const rule of opts.columnFormats ?? []) {
        if (!rule.forceText) continue;
        for (const idx of expandColumnRule(rule.columns, keepHeaders.length)) {
          forceTextTargets.add(idx);
        }
      }

      // The universal "no headers" rule strips the source's header row by default.
      // A preset can opt back in by setting keepHeader, in which case we write
      // a header row using each preset's `headerLabels` override (falling back to
      // the preset's match name).
      const outputRows: unknown[][] = [];
      const outputFormats: string[][] = [];
      if (opts.keepHeader) {
        const labels = keepHeaders.map((src, i) => opts.headerLabels?.[i] ?? src);
        outputRows.push(labels);
        outputFormats.push(keepHeaders.map(() => "General"));
      }
      for (let r = dataStart; r < data.length; r++) {
        const srcRow = data[r];
        const srcTextRow = sourceText[r] ?? [];
        const srcFmtRow = sourceFormats[r] ?? [];
        const srcTypeRow = sourceValueTypes[r] ?? [];
        outputRows.push(
          sourceCols.map((idx, targetIdx) => {
            if (idx === -1) return "";
            // forceText columns: ship the displayed string ("1 May") rather than
            // the underlying value, so a date-serial source doesn't ride out to
            // the clipboard as a real date and get pill-ified by Canva/Slides.
            if (forceTextTargets.has(targetIdx)) return srcTextRow[idx] ?? "";
            return srcRow[idx];
          })
        );
        outputFormats.push(
          sourceCols.map((idx, targetIdx) => {
            if (idx === -1) return "General";
            if (forceTextTargets.has(targetIdx)) return "@";
            // Force text format on cells whose source value is a string (e.g.,
            // the Birthdays report stores Birth Date as "6 May"). Without this,
            // Excel auto-parses the string as a date on writeback and the
            // inherited "General" format then displays it as a serial number.
            if (srcTypeRow[idx] === "String") return "@";
            return srcFmtRow[idx] ?? "General";
          })
        );
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

      // Phase 1: pre-format the target cells before any value is written.
      // Critical for cells we want as text ("@") — if we set values and "@"
      // in the same sync, Excel's auto-parser runs against the value before
      // the format takes effect, turning "6 May" into a date serial. Setting
      // numberFormat in its own sync first locks the cells as text up front.
      // This also installs the inherited per-cell source numberFormat so
      // columns without an explicit override (e.g., the % column on Checks)
      // keep their source display (4% instead of 0.04).
      target.numberFormat = outputFormats as Excel.Range["numberFormat"];
      await context.sync();

      // Phase 2: write values and apply the rest of the formatting. Per-rule
      // numberFormat in applyColumnFormats overrides phase 1 for currency etc.
      target.values = outputRows as Excel.Range["values"];

      applyUniversalFormat(target);
      applyColumnFormats(sheet, outputRows.length, keepHeaders.length, opts.columnFormats);

      if (opts.keepHeader) {
        // Bold the header row (only present when keepHeader is true).
        const headerRange = sheet.getRangeByIndexes(0, 0, 1, keepHeaders.length);
        headerRange.format.font.bold = true;
      }

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
