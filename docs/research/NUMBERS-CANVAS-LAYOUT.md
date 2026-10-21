# Native Numbers canvas layout

The supplied `Untitled 2.numbers` is retained byte-for-byte as
`tests/fixtures/numbers-multiple-sheets-charts.numbers` (SHA-256
`05af58cd0ae6af8df8cd46ab65e88b8060aacc2ce8c5c4d59baead2cd6b28566`).

## Native reference

Read directly from the original file open in Numbers on 2026-10-21. Coordinates
are document points; the adapter converts points to CSS pixels by 96/72.

| Object | x | y | width | height |
| --- | ---: | ---: | ---: | ---: |
| Monthly Budget text box | 0 | 0 | 503.3 | 42 |
| Budget instructions | 0 | 44 | 503 | 48 |
| Actual Summary doughnut plot | 48 | 134 | 153.5 | 153.5 |
| Budget vs. Actual column plot | 323.8 | 135 | 179.2 | 152 |
| Summary by Category table body | 0 | 440 | 503 | 226.6 |
| Transactions table body | 0 | 106 | 503.1 | 546.4 |

The former adapter moved each table to (0, 0), omitted floating text and the
Type 25 doughnut, and placed legends beside plots instead of at their saved
frames. Enlarging the viewport fixed clipping but did not fix those positions.

## Semantic boundary and reuse

- A sheet with native table geometry is a canvas, with one unit per native sheet.
  Unit rows/columns are zero because independent floating tables do not share
  one indexed grid. Cells keep their addresses and table-qualified stable IDs.
- Tables use their model's default row/column sizes plus explicit header overrides,
  and their own canvas origins. Auto-sized table names precede the table body.
  Pages/Keynote fixed-frame tables normalize sizes to their frame; that operation
  is deliberately not substituted for Numbers table layout.
- Geometry-free tile tables retain the existing grid adapter. No canvas origin
  is invented when the source lacks one.
- Ordinary TSD/TSWP/TSCH objects reuse the existing iWork emitters. Unbounded
  canvas text preserves its frame instead of clamping against a zero-size page.
- Doughnuts use the shared ellipse Bezier primitive, the file's inner-radius ratio
  and start angle. Legend frames and legend paragraph styles are read separately
  from dataset-name styles. Outside percentage labels use no-wrap text with zero
  insets. Native iWork drawable source kinds share one decoder.
- A malformed local drawable retains an omission diagnostic while other objects
  and sheets continue. An unindexed canvas has no formula viewport roots; printing
  retains its full canvas as a fragment rather than an empty cell range.

Source-field cross-checks:
[TSCH chart and legend references](https://github.com/m-tymchyk/apple-iwork-protos/blob/master/TSCHArchives.proto),
[TSCH generated property fields](https://github.com/m-tymchyk/apple-iwork-protos/blob/master/TSCHArchives.GEN.proto),
[TST table model fields](https://github.com/m-tymchyk/apple-iwork-protos/blob/master/TSTArchives.proto).
The current file and native application, rather than those historical schemas,
are the visual reference. Currency/date display formats, exact font-dependent
row auto-sizing and some chart styling remain approximate.

## Reproduction

```sh
node --test tests/spreadsheet.test.mjs tests/protocol.test.mjs tests/display-list.test.mjs tests/viewer-component.test.mjs
SHEET_DRAWING_FILE='/Users/wjfree/Downloads/Untitled 2.numbers' node scripts/test-sheet-drawing-extents.mjs
```

The browser regression runs Chromium and WebKit, validates native geometry,
compares tiled and direct pixels, then checks 200% zoom, edge scrolling and sheet
switching. It rejects an original file whose bytes differ from the retained case.
`SHEET_DRAWING_BASELINE_WASM` can substitute a pre-fix pack; both browsers fail
because only one chart is present and native captions/text frames are absent.
Native screenshots, before/after Viewer screenshots and local test logs are kept
under `output/playwright/` and are not embedded as document rendering resources.
