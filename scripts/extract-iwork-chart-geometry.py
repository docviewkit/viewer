"""Extract the pinned Charting Basics geometry oracle from an Apple Numbers PDF.

Usage: python3 scripts/extract-iwork-chart-geometry.py <native-reference-folder>
Requires PyMuPDF (the recorded capture used version 2.0). This is optional test
maintenance tooling; neither PDF conversion nor this library enters the SDK.
"""
import hashlib
import json
import sys
from pathlib import Path

import pymupdf

folder = Path(sys.argv[1])
reference = json.loads((folder / "reference.json").read_text())
fixture = Path("tests/fixtures/iwork-charting-basics.numbers")
pdf = folder / "reference.pdf"
assert reference["source"]["file"] == fixture.name
assert reference["referenceApplication"]["name"] == "Numbers"
assert reference["oracleSuite"] == "apple-iwork"
fixture_hash = hashlib.sha256(fixture.read_bytes()).hexdigest()
pdf_hash = hashlib.sha256(pdf.read_bytes()).hexdigest()
assert reference["source"]["sha256"] == fixture_hash
assert reference["pdfs"][0]["sha256"] == pdf_hash

# This fixture's native print layout puts the unchanged sheet canvas at (72,72)
# PDF points. Fixing that transform explicitly avoids comparing a print page to
# a worksheet viewport or letting automatic image alignment hide layout errors.
def css_point(point):
    return [(point.x - 72) * 4 / 3, (point.y - 72) * 4 / 3]


result = {
    "schemaVersion": 1,
    "fixture": fixture.name,
    "fixtureSha256": fixture_hash,
    "pdfSha256": pdf_hash,
    "extractorVersion": pymupdf.VersionBind,
    "method": "Apple Numbers PDF export; PyMuPDF get_drawings; sheet origin is PDF (72,72) points; points converted to CSS at 4/3",
    "radar": [],
    "bubbles": [],
    "dualMarkers": [],
}
with pymupdf.open(pdf) as document:
    assert len(document) == 8
    for drawing in document[2].get_drawings():
        if drawing["fill_opacity"] == 0.5 and len(drawing["items"]) == 5:
            color = drawing["fill"]
            series = 0 if color[2] > 0.9 else 1 if color[1] > 0.7 else 2
            result["radar"].append({
                "series": series,
                "points": [css_point(item[1]) for item in drawing["items"]],
                "opacity": 0.5,
            })
    for drawing in document[7].get_drawings():
        color, rect = drawing["fill"], drawing["rect"]
        if color and color[2] > 0.9 and color[0] < 0.1 and rect.width > 10 and rect.height > 10:
            result["bubbles"].append({
                "center": css_point(rect.tl + (rect.br - rect.tl) * 0.5),
                "radius": (rect.width + rect.height) / 3,
            })
    for drawing in document[5].get_drawings():
        color, rect = drawing["fill"], drawing["rect"]
        if color and min(color) > 0.99 and 10 < rect.width < 20 and rect.y0 > 400:
            result["dualMarkers"].append(css_point(rect.tl + (rect.br - rect.tl) * 0.5))
assert [len(result[key]) for key in ["radar", "bubbles", "dualMarkers"]] == [3, 9, 12]
print(json.dumps(result, indent=2))
