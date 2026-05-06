import { test } from "node:test";
import assert from "node:assert/strict";
import { normalizeHeader, detectHeaderRow, resolveColumnIndex, expandColumnRule } from "./headers";

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

test("expandColumnRule: single index returns one-element array", () => {
  assert.deepEqual(expandColumnRule(2, 5), [2]);
});

test("expandColumnRule: inclusive range expands all indices", () => {
  assert.deepEqual(expandColumnRule([1, 4], 7), [1, 2, 3, 4]);
});

test("expandColumnRule: clamps the end of a range to columnCount-1", () => {
  assert.deepEqual(expandColumnRule([1, 10], 5), [1, 2, 3, 4]);
});

test("expandColumnRule: clamps a single index past the end to empty", () => {
  assert.deepEqual(expandColumnRule(7, 3), []);
});

test("expandColumnRule: negative single index is empty", () => {
  assert.deepEqual(expandColumnRule(-1, 5), []);
});

test("expandColumnRule: range with start > columnCount is empty", () => {
  assert.deepEqual(expandColumnRule([10, 20], 5), []);
});

test("expandColumnRule: range with start > end is empty", () => {
  assert.deepEqual(expandColumnRule([4, 1], 7), []);
});

test("detectHeaderRow: row 1 with one extra non-empty cell stays at row 0 (Court of Sharing case)", () => {
  // The InTouch Court of Sharing report has a header row with an unlabeled
  // column (position 4) and the data row has a value there — so the data row
  // has exactly 1 more non-empty cell than the header row. The grouping-row
  // heuristic must NOT treat that as a multi-row header.
  const rows = [
    ["Date", "Rpt Title", "Rpt Unit", "Rpt Name", "", "Name", "Sem Qual Team Member", "Sem Recruiter Comm Earned"],
    ["APRIL-2026", "Seminar YTD Court of Sharing Report", "B235", "Brigitte Iglay", 1.0, "Gina Bateman", 1.0, 36.0],
  ];
  assert.deepEqual(detectHeaderRow(rows), { headerRowIdx: 0, groupingRowIdx: -1 });
});

test("detectHeaderRow: row 1 with two extra non-empty cells still picks row 1 (margin boundary)", () => {
  const rows = [
    ["", "", "", "Group A", "", "", ""],          // 1 non-empty
    ["A", "B", "C"],                                // 3 non-empty (margin=2)
  ];
  assert.deepEqual(detectHeaderRow(rows), { headerRowIdx: 1, groupingRowIdx: 0 });
});
