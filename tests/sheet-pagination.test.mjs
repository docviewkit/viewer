import assert from "node:assert/strict";
import test from "node:test";

import { sheetPrintPages } from "../dist/render.js";

test("paginates the authored Styles.xlsx Sheet1 on cell boundaries", () => {
  const unit = {
    type: "sheet",
    index: 0,
    id: "unit:0",
    name: "Sheet1",
    width: 1013,
    height: 1012,
    rows: 45,
    columns: 12,
    frozenRows: 0,
    frozenColumns: 0,
    frozenWidth: 0,
    frozenHeight: 0,
    rowAxis: {
      defaultSize: 22,
      spans: [
        { start: 1, end: 1, size: 23 },
        { start: 2, end: 3, size: 24 },
        { start: 4, end: 4, size: 23 },
        { start: 5, end: 5, size: 28 },
        { start: 6, end: 6, size: 32 },
      ],
    },
    columnAxis: {
      defaultSize: 69,
      spans: [
        { start: 0, end: 0, size: 126 },
        { start: 1, end: 1, size: 104 },
        { start: 2, end: 2, size: 74 },
        { start: 3, end: 3, size: 66 },
        { start: 4, end: 5, size: 86 },
        { start: 8, end: 8, size: 126 },
      ],
    },
    showGridLines: true,
    printSettings: {
      viewMode: "pageLayout",
      orientation: "portrait",
      fitToPage: false,
      margins: { left: .7, right: .7, top: .75, bottom: .75, header: .3, footer: .3 },
      differentOddEven: false,
      differentFirst: false,
    },
  };
  const pages = sheetPrintPages(unit);

  assert.deepEqual(pages.map(({ viewport }) => viewport), [
    { x: 0, y: 0, width: 680, height: 902 },
    { x: 0, y: 902, width: 680, height: 110 },
    { x: 680, y: 0, width: 333, height: 902 },
    { x: 680, y: 902, width: 333, height: 110 },
  ]);
  assert.deepEqual(pages[0].paper, { width: 816, height: 1056 });

  const fitted = sheetPrintPages({
    ...unit,
    printSettings: { ...unit.printSettings, fitToPage: true, fitToWidth: 1, fitToHeight: 2 },
  });
  assert.equal(new Set(fitted.map(({ viewport }) => viewport.x)).size, 1);
});
