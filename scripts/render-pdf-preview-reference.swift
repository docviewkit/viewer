import AppKit
import Foundation
import PDFKit

func fail(_ message: String) -> Never {
    FileHandle.standardError.write(Data("\(message)\n".utf8))
    exit(1)
}

func selectedPages(_ specification: String, count: Int) -> [Int] {
    if specification == "all" {
        return Array(0..<count)
    }
    var pages = Set<Int>()
    for part in specification.split(separator: ",") {
        let bounds = part.split(separator: "-", maxSplits: 1)
        guard let first = Int(bounds[0]), first >= 1 else {
            fail("Invalid page selection: \(part)")
        }
        let last = bounds.count == 2 ? Int(bounds[1]) : first
        guard let last, last >= first, last <= count else {
            fail("Page selection is outside the PDF: \(part)")
        }
        for page in first...last {
            pages.insert(page - 1)
        }
    }
    return pages.sorted()
}

guard CommandLine.arguments.count == 5 || CommandLine.arguments.count == 6 else {
    fail("Usage: render-pdf-preview-reference.swift <input.pdf> <output-directory> <scale> <all|pages> [dimensions.json]")
}
let input = URL(fileURLWithPath: CommandLine.arguments[1])
let output = URL(fileURLWithPath: CommandLine.arguments[2], isDirectory: true)
guard let scale = Double(CommandLine.arguments[3]), scale > 0, scale <= 4 else {
    fail("Scale must be in (0, 4]")
}
guard let document = PDFDocument(url: input), document.pageCount > 0 else {
    fail("PDFKit could not open \(input.path)")
}
try FileManager.default.createDirectory(at: output, withIntermediateDirectories: true)
var dimensions: [String: [String: Int]] = [:]
if CommandLine.arguments.count == 6 {
    let url = URL(fileURLWithPath: CommandLine.arguments[5])
    let value = try JSONSerialization.jsonObject(with: Data(contentsOf: url))
    guard let parsed = value as? [String: [String: Int]] else {
        fail("Dimensions JSON must map 1-based page numbers to width and height")
    }
    dimensions = parsed
}

let pages = selectedPages(CommandLine.arguments[4], count: document.pageCount)
for pageIndex in pages {
    guard let page = document.page(at: pageIndex) else {
        fail("PDFKit could not read page \(pageIndex + 1)")
    }
    let bounds = page.bounds(for: .cropBox)
    let requested = dimensions[String(pageIndex + 1)]
    let width = requested?["width"] ?? max(1, Int(ceil(bounds.width * scale)))
    let height = requested?["height"] ?? max(1, Int(ceil(bounds.height * scale)))
    guard width > 0, height > 0 else {
        fail("Dimensions for page \(pageIndex + 1) must be positive")
    }
    let scaleX = Double(width) / bounds.width
    let scaleY = Double(height) / bounds.height
    guard width * height <= 100_000_000,
          let bitmap = NSBitmapImageRep(
              bitmapDataPlanes: nil,
              pixelsWide: width,
              pixelsHigh: height,
              bitsPerSample: 8,
              samplesPerPixel: 4,
              hasAlpha: true,
              isPlanar: false,
              colorSpaceName: .deviceRGB,
              bytesPerRow: 0,
              bitsPerPixel: 0
          ),
          let graphics = NSGraphicsContext(bitmapImageRep: bitmap)
    else {
        fail("Could not allocate page \(pageIndex + 1) at \(width)x\(height)")
    }
    NSGraphicsContext.saveGraphicsState()
    NSGraphicsContext.current = graphics
    let context = graphics.cgContext
    context.setFillColor(NSColor.white.cgColor)
    context.fill(CGRect(x: 0, y: 0, width: width, height: height))
    context.scaleBy(x: scaleX, y: scaleY)
    context.translateBy(x: -bounds.minX, y: -bounds.minY)
    page.draw(with: .cropBox, to: context)
    NSGraphicsContext.restoreGraphicsState()
    guard let png = bitmap.representation(using: .png, properties: [:]) else {
        fail("Could not encode page \(pageIndex + 1)")
    }
    let filename = String(format: "page-%04d.png", pageIndex + 1)
    try png.write(to: output.appendingPathComponent(filename), options: .atomic)
}

print("Rendered \(pages.count) of \(document.pageCount) PDF pages with PDFKit/Preview at scale \(scale)")
