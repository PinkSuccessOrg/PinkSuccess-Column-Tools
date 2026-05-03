/* global Office, Excel */

type Operation = "keepInOrder" | "keepBySet" | "remove";

interface Preset {
  id: string;
  label: string;
  operation: Operation;
  headers: string[];
}

const PRESETS: Preset[] = [
  {
    id: "preset-customer", label: "Customer Export",
    operation: "keepInOrder", headers: ["Name", "Email", "Phone", "Status"]
  },
  {
    id: "preset-sales", label: "Sales Report",
    operation: "keepInOrder", headers: ["Date", "Rep", "Account", "Amount", "Stage"]
  },
  {
    id: "preset-strip", label: "Strip Internal Fields",
    operation: "remove", headers: ["InternalID", "Notes", "TempField"]
  },
];

function runPreset(p: Preset): Promise<void> {
  switch (p.operation) {
    case "keepInOrder": return keepColumnsInOrder(p.headers);
    case "keepBySet": return keepColumnsBySet(p.headers);
    case "remove": return removeColumnsByHeader(p.headers);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  applyTheme();

  const container = document.getElementById("presets")!;
  for (const p of PRESETS) {
    const btn = document.createElement("button");
    btn.className = "ms-Button";
    btn.id = p.id;
    btn.textContent = p.label;
    btn.addEventListener("click", () => void runPreset(p));
    container.appendChild(btn);
  }

  bindClick("run-keep-order", () => keepColumnsInOrder(parseCustom()));
  bindClick("run-keep-set", () => keepColumnsBySet(parseCustom()));
  bindClick("run-remove", () => removeColumnsByHeader(parseCustom()));
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

// ---------- The three mains ----------

async function keepColumnsInOrder(keepHeaders: string[]): Promise<void> {
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
      const headerRow = data[0].map((h) => String(h ?? ""));

      // Build new columns in requested order
      const newCols: unknown[][] = keepHeaders.map((wanted) => {
        const idx = headerRow.findIndex(
          (h) => h.toLowerCase() === wanted.toLowerCase()
        );
        if (idx === -1) {
          // Header missing in source — emit blank column with the header in row 0
          return [wanted, ...new Array<unknown>(data.length - 1).fill("")];
        }
        return data.map((row) => row[idx]);
      });

      // Transpose columns→rows for the target write
      const newValues: unknown[][] = data.map((_, r) =>
        newCols.map((col) => col[r])
      );

      used.clear();
      const target = sheet.getRangeByIndexes(
        0,
        0,
        newValues.length,
        newValues[0].length
      );
      target.values = newValues as Excel.Range["values"];
      await context.sync();

      setStatus(`Kept and reordered ${keepHeaders.length} column(s).`);
    });
  } catch (e) {
    setStatus(`Error: ${(e as Error).message}`);
  }
}

async function keepColumnsBySet(keepHeaders: string[]): Promise<void> {
  if (keepHeaders.length === 0) {
    setStatus("No headers given.");
    return;
  }
  const wanted = new Set(keepHeaders.map((h) => h.toLowerCase()));
  await deleteWhere((header) => !wanted.has(header.toLowerCase()), "Kept");
}

async function removeColumnsByHeader(removeHeaders: string[]): Promise<void> {
  if (removeHeaders.length === 0) {
    setStatus("No headers given.");
    return;
  }
  const drop = new Set(removeHeaders.map((h) => h.toLowerCase()));
  await deleteWhere((header) => drop.has(header.toLowerCase()), "Removed");
}

async function deleteWhere(
  shouldDelete: (header: string) => boolean,
  verb: string
): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();
      used.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const data = used.values as unknown[][];
      const headerRow = data[0].map((h) => String(h ?? ""));
      let deleted = 0;

      // Right-to-left so indices stay valid as we delete
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