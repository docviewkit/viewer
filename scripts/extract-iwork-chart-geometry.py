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

def chart_hue_columns(image):
    pixels = image.samples
    blue = [0] * image.width; red = [0] * image.width; red_top = [image.height] * image.width; green = [0] * image.width
    bounds = {"blue":[image.height,-1]}
    for y in range(image.height):
        for x in range(image.width):
            offset = (y*image.width+x)*image.n
            r,g,b = pixels[offset:offset+3]
            if b-r > 40 and b-g > 15:
                blue[x] += 1; bounds["blue"][0] = min(bounds["blue"][0],y); bounds["blue"][1] = y
            green[x] += g-r > 15 and g-b > 40
            if r-b > 40 and r-g > 15:
                red[x] += 1; red_top[x] = min(red_top[x],y)
    return blue,red,red_top,green,bounds

def chromatic_convex_coverage(mask, width, height):
    points = []
    for y in range(height):
        xs = [x for x in range(width) if mask[y*width+x]]
        if xs:
            points.extend([(xs[0],y),(xs[-1]+1,y),(xs[0],y+1),(xs[-1]+1,y+1)])
    points = sorted(set(points))
    def cross(a,b,c): return (b[0]-a[0])*(c[1]-a[1])-(b[1]-a[1])*(c[0]-a[0])
    def chain(points):
        result = []
        for p in points:
            while len(result)>1 and cross(result[-2],result[-1],p)<=0: result.pop()
            result.append(p)
        return result
    hull = chain(points)[:-1]+chain(reversed(points))[:-1]
    area = abs(sum(a[0]*b[1]-a[1]*b[0] for a,b in zip(hull,hull[1:]+hull[:1])))/2
    return sum(mask)/area if area else 0

folder = Path(sys.argv[1])
reference = json.loads((folder / "reference.json").read_text())
source_name = reference["source"]["file"]
if source_name in {"3d-six-default-lighting.key","3d-six-implicit-lighting.key"}:
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures")/source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder/"reference.pdf").read_bytes()).hexdigest()
    comparison = json.loads((Path("tests/fixtures/iwork-chart-oracles/3d-six-default-lighting/reference.json")).read_text())
    pairs=[]
    for i,name in enumerate(["column","line","area","pie","stacked-column","stacked-area"]):
        for j in [i,i+6]:
            assert hashlib.sha256((folder/reference["pages"][j]["file"]).read_bytes()).hexdigest()==reference["pages"][j]["sha256"]
        equal=reference["pages"][i]["sha256"]==comparison["pages"][i+6]["sha256"]
        assert equal == (name!="pie" or source_name=="3d-six-implicit-lighting.key"), (name,equal)
        pairs.append({"name":name,"implicit":i,"explicit":i+6,"nativeIdentical":equal})
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],"pdfSha256":reference["pdfs"][0]["sha256"],"method":"Real native controls: six original 3D chart types and independent duplicates after native UI Default -> Soft Fill -> Default. The original Pie is authored Glossy; the other five originals are explicit Default already. A separate file removes all lighting packages: its first six native pages are byte-identical to the six explicit Default duplicate pages, including Pie. Preserve authored Glossy separately. No full visual acceptance.","comparisonFixture":"3d-six-default-lighting.key","pairs":pairs},indent=2))
    sys.exit(0)

if source_name in {"axis-tick-boundaries.key", "axis-tick-horizontal.key", "axis-tick-logarithmic.key", "axis-tick-logarithmic-ranges.key"}:
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    units=[]
    with pymupdf.open(folder/"reference.pdf") as document:
        for i,page in enumerate(document):
            segments=[];ticks=[]
            for d in page.get_drawings():
                if d["color"] is not None and max(d["color"])<.01:
                    assert abs(d["width"]-1)<.001 and d["lineCap"]==(2,2,2)
                    for item in d["items"]:
                        assert item[0]=="l"
                        a,b=item[1:];line=[[p.x*4/3,p.y*4/3] for p in [a,b]];segments.append(line)
                        if abs(a.x-b.x)+abs(a.y-b.y)<10:ticks.append(line)
            assert len(segments)==({"axis-tick-horizontal.key":[11,11,15],"axis-tick-logarithmic.key":[25,25],"axis-tick-logarithmic-ranges.key":[16,34,25,16,25]}.get(source_name,[7,7,15,6,19]))[i],(i,len(segments))
            png=folder/reference["pages"][i]["file"]
            assert hashlib.sha256(png.read_bytes()).hexdigest()==reference["pages"][i]["sha256"]
            units.append({"unit":i,"categoryLabels":["Jan","Feb","Mar"],"labels":[v for v in page.get_text().splitlines() if v not in {"Jan","Feb","Mar"}],"blackSegments":segments,"tickSegments":ticks,"nativePageSha256":reference["pages"][i]["sha256"],"nativeFrame":{"x":0,"y":0,"width":reference["pages"][i]["width"],"height":reference["pages"][i]["height"],"scale":1}})
            if source_name == "axis-tick-logarithmic-ranges.key":
                bars=[]
                for drawing in page.get_drawings():
                    if drawing["fill"] and max(drawing["fill"])-min(drawing["fill"])>.1:
                        r=drawing["rect"];bars.append([r.x0*4/3,r.y0*4/3,r.width*4/3,r.height*4/3])
                assert len(bars)==[6,6,4,6,6][i]
                units[-1]["bars"]=bars
                units[-1]["gridSegments"]=[[[p.x*4/3,p.y*4/3] for p in item[1:]] for drawing in page.get_drawings() if drawing["color"] is not None and .1<min(drawing["color"])<.99 for item in drawing["items"] if item[0]=="l"]

    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],"pdfSha256":reference["pdfs"][0]["sha256"],"method":"Native Keynote Cartesian tick controls: inside, centered, outside with minor ticks/category marks, outside while value axis is hidden, and translated inside with two minor ticks. PDF black path endpoints, cap and width are extracted exactly. Native zero-length outside minor segments have no separate raster ink and remain contentless. Native 96-dpi full-page viewport fixes raster phase; focused binary black-ink patches do not replace the full visual policy.","units":units},indent=2))
    sys.exit(0)

if source_name in {"stack-sparse-negative.key","stack-cross-boundaries.key","stack-cross-boundaries-2d.key"}:
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    units = []
    with pymupdf.open(folder / "reference.pdf") as document:
        for i,page in enumerate(document):
            unit = {"unit":i,"categoryLabels":["April","May","June","July"] if source_name=="stack-cross-boundaries-2d.key" else ["A","B","C","D"]}
            labels = [v for v in page.get_text().splitlines() if v not in set(unit["categoryLabels"])]
            three_d = source_name=="stack-cross-boundaries.key" or (source_name=="stack-sparse-negative.key" and i==3)
            if three_d:
                if source_name=="stack-sparse-negative.key": labels = ["-10","-8","-5","-3","0","3","5","8","10"]
                elif i==1: labels = [f"{v:.2f}" for v in [-8,-4,0,4,8,12,16]]
                elif i==2: labels = [("-" if v<0 else "")+f"¥{abs(v):.2f}" for v in [-8,-4,0,4,8,12,16]]
                elif i==4: labels = ["-10","-5","0","5","10","15","20"]
                else: labels = ["-8","-4","0","4","8","12","16"]
                unit["threeDimensional"] = True
            png=folder/reference["pages"][i]["file"]
            assert hashlib.sha256(png.read_bytes()).hexdigest()==reference["pages"][i]["sha256"]
            unit["nativePageSha256"]=reference["pages"][i]["sha256"]
            if source_name=="stack-sparse-negative.key" and i in {0,1}:
                bars=[]
                for d in page.get_drawings():
                    if d["fill"] and max(d["fill"])-min(d["fill"])>.1:
                        r=d["rect"];bars.append([r.x0*4/3,r.y0*4/3,r.width*4/3,r.height*4/3])
                assert len(bars)==4
                unit["bars"]=bars
            if source_name=="stack-sparse-negative.key" and i==2:
                boundaries=[]
                for d in page.get_drawings():
                    if d["fill"] and max(d["fill"])-min(d["fill"])>.1:
                        points=[]
                        for item in d["items"]:
                            assert item[0]=="l"
                            points.extend([[p.x*4/3,p.y*4/3] for p in item[1:]])
                        points.append(points[0]);boundaries.append(points)
                assert len(boundaries)==2
                unit["areaBoundaries"]=boundaries
            unit["labels"]=labels;units.append(unit)
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],"pdfSha256":reference["pdfs"][0]["sha256"],"method":"Native Keynote PDF text and vectors for 2D; each 3D PNG and numeric label visually inspected, source and PNG hashes pinned. Sparse negative stack retains both signed halves and zero category valleys. Boundary controls cover explicit Number zero/two decimals, Currency two, major four and Auto zero. No full visual acceptance.","units":units},indent=2))
    sys.exit(0)

if source_name == "stack-opposite-signs.key":
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    units = []
    with pymupdf.open(folder / "reference.pdf") as document:
        for i,page in enumerate(document):
            labels = [v for v in page.get_text().splitlines() if v not in {"A","B","C","D"}]
            unit = {"unit":i,"categoryLabels":["A","B","C","D"]}
            if i>=3:
                labels = ["-20","-15","-10","-5","0","5","10","15","20"] if i==5 else ["-30","-22.5","-15","-7.5","0","7.5","15","22.5","30"]
                png = folder / reference["pages"][i]["file"]
                assert hashlib.sha256(png.read_bytes()).hexdigest() == reference["pages"][i]["sha256"]
                unit.update(threeDimensional=True,nativePageSha256=reference["pages"][i]["sha256"])
            if i in {0,1}:
                bars = []
                for d in page.get_drawings():
                    if d["fill"] and max(d["fill"])-min(d["fill"])>.1:
                        r = d["rect"]; bars.append([r.x0*4/3,r.y0*4/3,r.width*4/3,r.height*4/3])
                assert len(bars)==8,(i,len(bars))
                unit["bars"] = bars
            unit["labels"] = labels
            units.append(unit)
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],
                     "pdfSha256":reference["pdfs"][0]["sha256"],
                     "method":"Apple Keynote opposite signs within each category: Alpha alternates +/-20, Beta alternates -/+30. Native 2D PDF text and 16 bar rectangles; all three 3D PNG labels visually inspected and hash pinned. Column/bar have separate signed extents +/-30; area cumulative boundaries reach +/-20. This records the adapter boundary, not full visual acceptance.","units":units},indent=2))
    sys.exit(0)

if source_name == "stack-cross-axis.key":
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    units = []
    with pymupdf.open(folder / "reference.pdf") as document:
        for i,page in enumerate(document):
            labels = [v for v in page.get_text().splitlines() if v not in {"A","B","C","D"}]
            unit = {"unit":i,"categoryLabels":["A","B","C","D"]}
            if i%6>=3:
                labels = ["-16","-12","-8","-4","0","4","8","12","16"] if i<6 else ["-8","-4","0","4","8","12","16"]
                png = folder / reference["pages"][i]["file"]
                assert hashlib.sha256(png.read_bytes()).hexdigest() == reference["pages"][i]["sha256"]
                unit.update(threeDimensional=True,nativePageSha256=reference["pages"][i]["sha256"])
            if i%6 in {0,1}:
                bars = []
                for d in page.get_drawings():
                    if d["fill"] and max(d["fill"])-min(d["fill"])>.1:
                        r = d["rect"]; bars.append([r.x0*4/3,r.y0*4/3,r.width*4/3,r.height*4/3])
                assert len(bars)==8,(i,len(bars))
                unit["bars"] = bars
            unit["labels"] = labels
            units.append(unit)
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],
                     "pdfSha256":reference["pdfs"][0]["sha256"],
                     "method":"Apple Keynote six stacked chart families with symmetric +/-15 and asymmetric -5..15 totals. 2D native PDF complete numeric text and bar rectangles; all six 3D pages individually inspected and native PNG hashes pinned. Native bounds are +/-16 and -8..16 with step four, independently of SDK.","units":units},indent=2))
    sys.exit(0)

if source_name == "3d-signed-area.key":
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    labels = [["-10","-7.5","-5","-2.5","0","2.5","5","7.5","10"],
              ["-16","-12","-8","-4","0","4","8","12","16"],
              ["-10","-7.5","-5","-2.5","0"],
              ["-10","-7.5","-5","-2.5","0","2.5","5","7.5","10"],
              ["-10","-5","0","5","10","15","20"]]
    units = []
    for i,page in enumerate(reference["pages"]):
        png = folder / page["file"]
        assert hashlib.sha256(png.read_bytes()).hexdigest() == page["sha256"]
        image = pymupdf.Pixmap(str(png)); blue,_red,_top,green,bounds = chart_hue_columns(image)
        left = next(x for x,n in enumerate(blue) if n); right = max(x for x,n in enumerate(blue) if n)
        coverage = sum(blue)/((right-left+1)*(bounds["blue"][1]-bounds["blue"][0]+1))
        units.append({"unit":i,"nativePageSha256":page["sha256"],"threeDimensional":True,
                      "labels":labels[i],"categoryLabels":["A","B","C"] if i==2 else ["A","B","C","D"],
                      "blueBodyCoverage":coverage,"greenPresent":sum(green)>=16})
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],
                     "pdfSha256":reference["pdfs"][0]["sha256"],
                     "method":"Manual inspection of every native 3D numeric label, pinned native PNG; area hue coverage independently counted from native raster. Alternating signed bands have substantial front faces even when signed polygon area cancels. Body regression retains at least 90% of native hue/bounding-box coverage; full visual policy remains separate and unchanged.","units":units},indent=2))
    sys.exit(0)

if source_name == "iwork-chart-3d-sparse-area.key":
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    units = []
    if source_name.startswith("3d-default-lighting"):
        assert len(reference["pages"])==2 and reference["pages"][0]["sha256"]==reference["pages"][1]["sha256"]
    for index, page in enumerate(reference["pages"]):
        png = folder / page["file"]
        assert hashlib.sha256(png.read_bytes()).hexdigest() == page["sha256"]
        image = pymupdf.Pixmap(str(png))
        blue,red,red_top,_green,_bounds = chart_hue_columns(image)
        left = next(x for x,n in enumerate(blue) if n)
        right = max(x for x,n in enumerate(blue) if n)
        low,high = [round(left+(right-left)*t) for t in (.2,.45)]
        valley = min(blue[low:high])/max(blue)
        red_left = next(x for x,n in enumerate(red) if n)
        red_right = max(x for x,n in enumerate(red) if n)
        x0,x1 = [round(red_left+(red_right-red_left)*t) for t in (.05,.45)]
        points = [(x,red_top[x]) for x in range(x0,x1) if red[x]]
        mx = sum(x for x,y in points)/len(points); my = sum(y for x,y in points)/len(points)
        slope = sum((x-mx)*(y-my) for x,y in points)/sum((x-mx)**2 for x,y in points)
        straight_rms = (sum((y-my-slope*(x-mx))**2 for x,y in points)/len(points))**.5
        units.append({"unit":index,"nativePageSha256":page["sha256"],
                      "redPixels":sum(red),"blueValleyRatio":valley,
                      "relativeRedStart":(red_left-left)/(right-left),
                      "redPresent":sum(red)>=16,
                      **({"blueValleyNearZero":valley<.1} if index in {0,2} else {}),
                      **({"redStartsAtFirstCategory":(red_left-left)/(right-left)<.05,"redBandHasArea":sum(red)>=sum(blue)*.5,"redTopStraightRms":straight_rms,"redTopIsStraight":straight_rms<1} if index==2 else {})})
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],
                     "pdfSha256":reference["pdfs"][0]["sha256"],
                     "method":"Native conversion of the same sparse data: Alpha=[10,blank,30,20], Beta=[blank,20,blank,40]. Apple 3D line bridges blanks; 3D area uses zero; stacked area includes each category in both cumulative boundaries. Raster hue masks use blue B-R>40/B-G>15, red R-B>40/R-G>15. Area valley is min blue column occupancy over 20%-45% blue span divided by its maximum (<0.1); stacked red starts within 5% of the blue left bound and has at least half the blue pixel area (native bands have comparable integral area). The first three stacked upper points (10,20,30) are collinear: red top-edge regression over 5%-45% red span has RMS below one native pixel. These topology probes do not replace full visual acceptance.",
                     "units":units},indent=2))
    sys.exit(0)

if source_name == "iwork-chart-cross-axis.key":
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    units = []
    with pymupdf.open(folder / "reference.pdf") as document:
        for i, page in enumerate(document):
            unit = {"unit": i, "labels": [s for s in page.get_text().splitlines() if s not in {"A", "B", "C"}],
                    "method": "Complete native PDF axis text; category labels excluded"}
            if i >= 24:
                unit.update(labels=["-15", "-7.5", "0", "7.5", "15", "22.5", "30"],
                            method="Manual inspection of all seven visible labels in Apple native 3D PDF raster; no SDK-derived oracle",
                            nativePageSha256=reference["pages"][i]["sha256"], threeDimensional=True)
            units.append(unit)
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],
                     "pdfSha256":reference["pdfs"][0]["sha256"],"extractorVersion":pymupdf.VersionBind,
                     "method":"Apple Keynote native mixed-sign automatic-axis heldouts, decimal thresholds and four nonstacked 3D types", "units":units},indent=2))
    sys.exit(0)

if source_name == "iwork-chart-single-sided-types.key":
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    units = []
    with pymupdf.open(folder / "reference.pdf") as document:
        for i,page in enumerate(document):
            unit = {"unit":i,"labels":[s for s in page.get_text().splitlines() if s not in {"A","B","C"}],
                    "method":"Complete native PDF axis text; category labels excluded"}
            if i in {1,2,3,4,8,9,10}:
                unit.update(labels=["0","650","1300","1950","2600"] if i < 5 else ["0","750","1500","2250","3000"],
                            method="Manual inspection of all five visible labels in Apple native 3D PDF raster; no SDK-derived oracle",
                            nativePageSha256=reference["pages"][i]["sha256"],threeDimensional=True)
            units.append(unit)
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],
                     "pdfSha256":reference["pdfs"][0]["sha256"],"extractorVersion":pymupdf.VersionBind,
                     "method":"Apple Keynote add chart data API: nonstacked values 1000/1600/2500 versus two-series stacks with identical 1000/1600/2500 totals; native bounds differ 2600/3000", "units":units},indent=2))
    sys.exit(0)

if source_name in {"iwork-chart-3d-clipped.key", "iwork-chart-3d-series-gap.key", "3d-zero-area.key", "3d-flat-area.key", "3d-default-lighting.key", "3d-default-lighting-stack-source.key", "3d-light-off-unique.key", "3d-center-intensity-four.key"} or source_name.startswith("3d-world-diffuse-"):
    from collections import deque
    units = []
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    for index, page in enumerate(reference["pages"]):
        png = folder / page["file"]
        assert hashlib.sha256(png.read_bytes()).hexdigest() == page["sha256"]
        image = pymupdf.Pixmap(str(png))
        pixels = image.samples
        mask = bytearray(max(pixels[i:i+3])-min(pixels[i:i+3]) > 40
                         for i in range(0,len(pixels),image.n))
        convex_coverage = chromatic_convex_coverage(mask,image.width,image.height) if source_name=="3d-flat-area.key" and index==0 else None
        colored_pixels = sum(mask)
        crop = None
        if source_name == "iwork-chart-3d-series-gap.key":
            colored_x = [i % image.width for i, value in enumerate(mask) if value]
            left, right = min(colored_x), max(colored_x)
            crop = [round(left + (right-left)*.8), round(left + (right-left)*.9)]
            for i in range(len(mask)):
                if not crop[0] <= i % image.width < crop[1]: mask[i] = 0
        components = []
        for start in range(len(mask)):
            if not mask[start]: continue
            mask[start] = 0
            queue = deque([start]); area = 0
            while queue:
                p = queue.popleft(); area += 1
                x,y = p % image.width,p // image.width
                for dy in (-1,0,1):
                    for dx in (-1,0,1):
                        if 0 <= x+dx < image.width and 0 <= y+dy < image.height:
                            q = p+dy*image.width+dx
                            if mask[q]: mask[q]=0; queue.append(q)
            if area >= 16: components.append(area)
        face_colors = []
        if source_name.startswith("3d-default-lighting") or source_name in {"3d-light-off-unique.key","3d-center-intensity-four.key"} or source_name.startswith("3d-world-diffuse-"):
            # Interior patches on the first visible native bar, independently
            # inspected in Apple's PNG; avoid bevels, highlights and silhouettes.
            for name,x,y in [("top",300,348),("front",290,369),("right",456,375)]:
                patch = [pixels[(py*image.width+px)*image.n:(py*image.width+px)*image.n+3]
                         for py in range(y-2,y+3) for px in range(x-2,x+3)]
                face_colors.append({"face":name,"nativePatch":[x-2,y-2,5,5],
                                    "rgb":[sorted(c[k] for c in patch)[12] for k in range(3)]})
        units.append({"unit": index, "nativePageSha256": page["sha256"],
                      "coloredComponents": len(components), "componentAreas": sorted(components), **({"faceColors":face_colors} if face_colors else {}), **({"convexCoverage":convex_coverage,"convexSilhouette":convex_coverage>=.995} if convex_coverage is not None else {}), **({"relativeChromaticXCrop":[.8,.9]} if crop else {}), **({"coloredPixels":colored_pixels,"bluePresent":sum(chart_hue_columns(image)[0])>=16,"greenPresent":sum(chart_hue_columns(image)[3])>=16} if source_name in {"3d-zero-area.key","3d-light-off-unique.key"} else {})})
    print(json.dumps({"fixture": source_name,"fixtureSha256":reference["source"]["sha256"],
                     "pdfSha256":reference["pdfs"][0]["sha256"],
                     "method":("Native implicit Default and explicit Default (after Soft Fill roundtrip) have byte-identical Apple PNGs. Bar and stacked-bar files independently preserve all data and camera controls. This proves default-light semantics, not full material or camera fidelity." if source_name.startswith("3d-default-lighting") else "Native world-space signed-axis directional light, unit intensity, white diffuse and zero specular; independent light name prevents Native transform reuse. Native face-interior RGB isolates normal orientation from point position and specular." if source_name.startswith("3d-world-diffuse-") else "Native single Center point light at intensity 4: sample end-cap diffuse saturation against Apple RGB, with point-position, bevel and specular still unaccepted." if source_name=="3d-center-intensity-four.key" else "Native explicit all-disabled four-light rig: visible achromatic dark bodies and zero chromatic regions. Unit 0 retains implicit Default. This does not prove camera or bevel fidelity." if source_name=="3d-light-off-unique.key" else "Native constant 3D area is a convex extruded rectangle. Chromatic pixel cells are compared with their convex hull: at least 99.5% coverage, allowing only raster boundary quantization. Native line remains a single independent strip. This does not replace the full visual policy." if source_name=="3d-flat-area.key" else "Native Apple 3D all-zero area is empty, zero stacked layers preserve the nonzero volume in either series order, zero line remains visible, and two adjacent zero area values split the positive regions. Hue and connected-region probes are independent of SDK output." if source_name=="3d-zero-area.key" else "Apple PDF native 3D raster, Poppler 96 dpi: 8-connected chromatic regions (RGB spread >40, area >=16 pixels). A=5 below manual 10..30 produces no visible bar; B=20/C=40 produce two bars." if source_name == "iwork-chart-3d-clipped.key" else
                               "Apple PDF native raster: count 8-connected chromatic regions (RGB spread >40, area >=16) in the 80%-90% horizontal span of colored geometry, away from the crossing on the left. Identical line series touch at 0% and separate at 100%/200%; stacked area layers touch. Ordinary area silhouettes overlap in projection, so this count does not prove their depth gap."),
                     "units": units},indent=2))
    sys.exit(0)

if source_name in {"iwork-chart-radar-labels.key", "iwork-chart-radar-value-axis.key"}:
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256((folder / "reference.pdf").read_bytes()).hexdigest()
    with pymupdf.open(folder / "reference.pdf") as document:
        units = []
        for i,page in enumerate(document):
            spokes = []
            axis_lines = []
            for d in page.get_drawings():
                if d["type"] == "s" and d["color"] and d["color"] != (0,0,0):
                    for item in d["items"]:
                        if item[0] == "l" and abs(item[1].x-item[2].x)<.1 and abs(item[1].y-item[2].y)>20:
                            axis_lines.append({"points":[[p.x*4/3,p.y*4/3] for p in item[1:]],"width":d["width"]*4/3})
                if d["type"] == "s" and d["color"] == (0,0,0):
                    for item in d["items"]:
                        if item[0] == "l":
                            spokes.append({"points":[[p.x*4/3,p.y*4/3] for p in item[1:]],"width":d["width"]*4/3})
            units.append({"unit":i,"labels":[s for s in page.get_text().splitlines() if s != "Manual Axis 5 to 50"],"spokes":spokes,"valueAxisLines":axis_lines})
    print(json.dumps({"fixture":source_name,"fixtureSha256":reference["source"]["sha256"],
                     "pdfSha256":reference["pdfs"][0]["sha256"],"extractorVersion":pymupdf.VersionBind,
                     "method":"Complete Apple PDF axis label sets, including category labels; title excluded", "units":units},indent=2))
    sys.exit(0)

if source_name == "iwork-chart-radar-grid.key":
    pdf = folder / "reference.pdf"
    assert reference["source"]["sha256"] == hashlib.sha256((Path("tests/fixtures") / source_name).read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256(pdf.read_bytes()).hexdigest()
    units = []
    fills = []
    with pymupdf.open(pdf) as document:
        for index, page in enumerate(document):
            grids = []
            probes = []
            drawings = page.get_drawings()
            fill = next(d for d in drawings if d["fill"] and max(d["fill"])-min(d["fill"])>.2)
            fills.append([fill["rect"].x0*4/3,fill["rect"].y0*4/3])
            for drawing in drawings:
                if drawing["type"] != "s" or len(drawing["items"]) < 3:
                    continue
                boundary = []
                for item in drawing["items"]:
                    if item[0] == "l":
                        boundary.extend([[p.x*4/3, p.y*4/3] for p in item[1:]])
                    elif item[0] == "c":
                        for step in range(33):
                            t = step/32
                            boundary.append([sum(weight*getattr(p, axis) for weight, p in
                                zip([(1-t)**3,3*(1-t)**2*t,3*(1-t)*t*t,t**3], item[1:]))*4/3
                                for axis in ["x", "y"]])
                first = drawing["items"][0]
                probes.append([((first[1].x+first[2].x)/2 if first[0]=="l" else first[1].x)*4/3,
                               ((first[1].y+first[2].y)/2 if first[0]=="l" else first[1].y)*4/3])
                grids.append({"boundary": boundary, "width": drawing["width"]*4/3,
                              "color": list(drawing["color"])})
            units.append({"unit": index, "grids": grids, "gridProbes": probes})
    # The identical native series establishes the page-0 to page-2 translation.
    units[2]["gridProbes"] = [[p[k]+fills[2][k]-fills[0][k] for k in range(2)]
                               for p in units[0]["gridProbes"]]
    print(json.dumps({"fixture": source_name, "fixtureSha256": reference["source"]["sha256"],
                     "pdfSha256": reference["pdfs"][0]["sha256"], "extractorVersion": pymupdf.VersionBind,
                     "method": "Apple PDF unfilled paths with at least three segments; cubic boundaries sampled directly in PDF points times 4/3; no fitted alignment", "units": units}, indent=2))
    sys.exit(0)

if source_name in {"iwork-chart-formatting.key", "iwork-chart-auto-axis.key", "iwork-chart-cross-formats.key", "one-sided-currency.key", "one-sided-number-two.key", "auto-count-heldout.key", "auto-count.key", "iwork-chart-label-boundaries.key", "iwork-chart-edge-cases.key", "chart-original.numbers", "iwork-static-matrix-floating.pages", "iwork-static-matrix-inline.pages"}:
    fixture = Path("tests/fixtures") / source_name
    pdf = folder / "reference.pdf"
    assert reference["oracleSuite"] == "apple-iwork"
    assert reference["source"]["sha256"] == hashlib.sha256(fixture.read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256(pdf.read_bytes()).hexdigest()
    result = {"fixture": source_name, "fixtureSha256": reference["source"]["sha256"],
              "pdfSha256": reference["pdfs"][0]["sha256"], "extractorVersion": pymupdf.VersionBind,
              "method": "Apple native PDF vectors and text; PDF points times 4/3 without fitted alignment; original Numbers sheet has a pinned (72,72) print translation", "units": []}
    with pymupdf.open(pdf) as document:
        for index, page in enumerate(document):
            if source_name == "iwork-chart-label-boundaries.key":
                labels = [s for s in page.get_text().splitlines()
                          if s not in {"Manual Axis 5 to 50", "A", "B", "C"}]
                unit = {"unit": index, "labels": labels,
                        "method": "Native PDF text; complete axis-label set"}
                if 3 <= index <= 4:
                    unit.update(labels=["16", "28", "39", "50"],
                                method="Manual observation of Apple native PDF 3D raster: four visible labels and no minimum 5; title excluded",
                                nativePageSha256=reference["pages"][index]["sha256"])
                if index < 2:
                    unit["bars"] = [[v*4/3 for v in (d["rect"].x0,d["rect"].y0,d["rect"].width,d["rect"].height)]
                        for d in page.get_drawings() if d["fill"] and max(d["fill"])-min(d["fill"])>.2]
                result["units"].append(unit)
                continue
            if source_name in {"iwork-chart-formatting.key", "iwork-chart-auto-axis.key", "iwork-chart-cross-formats.key", "one-sided-currency.key", "one-sided-number-two.key", "auto-count-heldout.key", "auto-count.key"}:
                # A label can use multiple fonts (for example a CJK prefix).
                lines = ["".join(s["text"] for s in line["spans"])
                         for block in page.get_text("dict")["blocks"] if "lines" in block
                         for line in block["lines"]]
                unit = {"unit": index, "labels": [s for s in lines
                    if s not in {"Currency", "Scientific", "Prefix suffix", "Fractions", "A", "B", "C"}]}
                if source_name in {"auto-count-heldout.key", "auto-count.key"}:
                    unit["nativeBarCount"] = sum(1 for d in page.get_drawings()
                        if d["fill"] and max(d["fill"])-min(d["fill"]) > .2)
                result["units"].append(unit)
                continue
            shapes = []
            for drawing in page.get_drawings():
                color = drawing.get("fill")
                if not color or max(color)-min(color) < .2:
                    continue
                if source_name == "chart-original.numbers" and not (
                    abs(color[0]-.35686275) < .0001 and abs(color[1]-.607843) < .0001):
                    continue
                rect = drawing["rect"]
                offset = 72 if source_name == "chart-original.numbers" else 0
                shape = {"bounds": [(rect.x0-offset)*4/3, (rect.y0-offset)*4/3,
                                     rect.width*4/3, rect.height*4/3], "fill": list(color)}
                if source_name == "iwork-chart-edge-cases.key" and index == 0:
                    boundary = []
                    for item in drawing["items"]:
                        if item[0] == "l":
                            boundary.extend([[p.x*4/3, p.y*4/3] for p in item[1:]])
                        elif item[0] == "c":
                            a, b, c, d = item[1:]
                            for step in range(33):
                                t = step/32
                                boundary.append([sum(weight*getattr(p, axis) for weight, p in
                                    zip([(1-t)**3, 3*(1-t)**2*t, 3*(1-t)*t*t, t**3], [a,b,c,d]))*4/3
                                    for axis in ["x", "y"]])
                    shape["boundary"] = boundary
                shapes.append(shape)
            result["units"].append({"unit": index, "labels": page.get_text().splitlines(), "shapes": shapes})
    if source_name == "chart-original.numbers":
        assert len(result["units"]) == 1 and len(result["units"][0]["shapes"]) == 2
        result["bars"] = sorted(s["bounds"] for s in result["units"][0]["shapes"])
    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0)

if reference["source"]["file"] == "iwork-chart-boundaries.key":
    fixture = Path("tests/fixtures/iwork-chart-boundaries.key")
    pdf = folder / "reference.pdf"
    assert reference["referenceApplication"]["name"] == "Keynote"
    assert reference["source"]["sha256"] == hashlib.sha256(fixture.read_bytes()).hexdigest()
    assert reference["pdfs"][0]["sha256"] == hashlib.sha256(pdf.read_bytes()).hexdigest()
    result = {"fixture": fixture.name, "fixtureSha256": reference["source"]["sha256"],
              "pdfSha256": reference["pdfs"][0]["sha256"], "extractorVersion": pymupdf.VersionBind,
              "method": "Apple Keynote 15.4 PDF get_drawings: colored paths and rectangles; CSS coordinates are PDF points times 4/3; no fitting or alignment", "units": []}
    with pymupdf.open(pdf) as document:
        assert len(document) == 8
        for index, page in enumerate(document):
            drawings = []
            for drawing in page.get_drawings():
                color = drawing["fill"] or drawing["color"]
                if not color or max(color) - min(color) < .2:
                    continue
                points = []
                for item in drawing["items"]:
                    if item[0] == "l":
                        points.extend([[item[1].x*4/3, item[1].y*4/3],
                                       [item[2].x*4/3, item[2].y*4/3]])
                rect = drawing["rect"]
                drawings.append({"color": list(color), "filled": drawing["fill"] is not None,
                                 "opacity": drawing["fill_opacity"],
                                 "bounds": [v*4/3 for v in (rect.x0, rect.y0, rect.width, rect.height)],
                                 "points": points})
            result["units"].append({"unit": index, "labels": page.get_text().splitlines(),
                                    "drawings": drawings})
    print(json.dumps(result, indent=2))
    sys.exit(0)

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
