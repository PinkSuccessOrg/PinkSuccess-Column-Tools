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
