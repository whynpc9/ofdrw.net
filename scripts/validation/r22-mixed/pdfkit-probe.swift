// Read-only native PDFKit API probe. Does not launch Preview or use clipboard.
import Foundation
import PDFKit
import CryptoKit

guard CommandLine.arguments.count == 3 else { fatalError("matrix.json output.json") }
let manifestData = try Data(contentsOf: URL(fileURLWithPath: CommandLine.arguments[1]))
let manifest = try JSONSerialization.jsonObject(with: manifestData) as! [String: Any]
let cases = manifest["cases"] as! [[String: Any]]
var rows = [[String: Any]]()
for item in cases {
    let path = item["path"] as! String
    let data = try Data(contentsOf: URL(fileURLWithPath: path))
    let hash = SHA256.hash(data: data).map { String(format: "%02x", $0) }.joined()
    guard hash == item["sha256"] as? String, let document = PDFDocument(data: data),
          document.pageCount == 1 else { fatalError("Hash/open/page failure: \(path)") }
    let selection = document.selectionForEntireDocument?.string ?? ""
    let expected = item["expected"] as! String
    let acceptedSelection = item["expectedSelection"] as! [String]
    var finds = [[String: Any]]()
    for word in item["expectedWords"] as! [[String: Any]] {
        let logical = word["text"] as! String
        for query in [logical, String(logical.reversed())] {
            let matches = document.findString(query, withOptions: [])
            let selections: [[String: Any]] = matches.map { match in
                let bounds: [[String: Any]] = match.pages.map { page in
                    let rect = match.bounds(for: page)
                    return ["page": document.index(for: page) + 1, "left": rect.minX,
                            "bottom": rect.minY, "right": rect.maxX, "top": rect.maxY,
                            "positiveArea": rect.width > 0 && rect.height > 0]
                }
                return ["text": match.string ?? "", "bounds": bounds]
            }
            finds.append(["query": query, "logical": query == logical,
                          "expectedCount": query == logical ? word["count"] as! Int : 0,
                          "count": matches.count, "matches": selections])
        }
    }
    rows.append(["name": item["name"]!, "path": path, "sha256": hash,
                 "document": document.string ?? "", "selection": selection,
                 "selectionCodepoints": selection.unicodeScalars.map { String(format: "U+%04X", $0.value) },
                 "exactSource": selection == expected,
                 "matchesDeclaredLineBreaks": acceptedSelection.contains(selection),
                 "finds": finds, "preview": "NOT TESTED: API-only result"])
}
let receipt: [String: Any] = ["os": ProcessInfo.processInfo.operatingSystemVersionString,
                             "pdfkitVersion": Bundle(for: PDFDocument.self).infoDictionary?["CFBundleVersion"] ?? "unknown",
                             "rows": rows]
let output = try JSONSerialization.data(withJSONObject: receipt, options: [.prettyPrinted, .sortedKeys])
try output.write(to: URL(fileURLWithPath: CommandLine.arguments[2]))
print("Recorded \(rows.count) native PDFKit reads; no Preview or clipboard access.")
