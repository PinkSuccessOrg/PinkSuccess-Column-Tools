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
  // Optional per-column display labels. Sparse — index `i` falls back to
  // `headers[i]` when the entry is missing or undefined. Lets a preset
  // abbreviate a verbose source column for the newsletter output.
  headerLabels?: (string | undefined)[];
  // Default false: the universal "no headers" rule strips the header row from
  // the output. Set true on a per-preset basis when the newsletter wants the
  // column titles preserved (e.g., OT Star).
  keepHeader?: boolean;
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
    headerLabels: [undefined, "Whsl $ + TB"],
    keepHeader: true,
    columnFormats: [
      { columns: 0, horizontalAlignment: "Left" },
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
    columnFormats: [
      { columns: 1, horizontalAlignment: "Left" },
    ],
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
