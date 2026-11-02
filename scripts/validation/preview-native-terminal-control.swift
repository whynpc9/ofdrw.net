// Minimal extraction of the existing native-font-word-controls/produce.swift
// Helvetica cases. Diagnostic only; does not use or change OFDRW.
// This extraction was NOT executed for the 2026-11-02 provenance review.
// Run on macOS: swift preview-native-terminal-control.swift /absolute/output
import Foundation
import CoreGraphics
import CoreText

precondition(CommandLine.arguments.count == 2, "Expected output directory")
let output = URL(fileURLWithPath: CommandLine.arguments[1], isDirectory: true)
try FileManager.default.createDirectory(at: output, withIntermediateDirectories: true)
let systemFont = CTFontCreateWithName("Helvetica" as CFString, 10.001, nil)
let font = CTFontCreateWithGraphicsFont(CTFontCopyGraphicsFont(systemFont, nil), 10.001, nil, nil)
precondition(CTFontCopyPostScriptName(font) as String == "Helvetica")
for (name, text) in [("helvetica-date", "2026-01-01"), ("helvetica-word", "Alpha")] {
    let url = output.appendingPathComponent(name + ".pdf")
    precondition(!FileManager.default.fileExists(atPath: url.path), "Do not replace frozen evidence")
    let line = CTLineCreateWithAttributedString(NSAttributedString(string: text, attributes: [
        NSAttributedString.Key(kCTFontAttributeName as String): font,
        NSAttributedString.Key(kCTForegroundColorAttributeName as String): CGColor(gray: 0, alpha: 1)
    ]))
    for run in CTLineGetGlyphRuns(line) as! [CTRun] {
        let actual = (CTRunGetAttributes(run) as NSDictionary)[kCTFontAttributeName] as! CTFont
        precondition(CTFontCopyPostScriptName(actual) as String == "Helvetica")
    }
    var box = CGRect(x: 0, y: 0, width: 595.275590551, height: 841.88976378)
    let context = CGContext(url as CFURL, mediaBox: &box, nil)!
    context.beginPDFPage(nil)
    context.textMatrix = .identity
    context.textPosition = CGPoint(x: 472.1472745, y: 686)
    context.setTextDrawingMode(.fill)
    CTLineDraw(line, context)
    context.endPDFPage()
    context.closePDF()
    print(url.path)
}
