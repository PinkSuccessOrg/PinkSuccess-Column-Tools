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
