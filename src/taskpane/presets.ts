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
