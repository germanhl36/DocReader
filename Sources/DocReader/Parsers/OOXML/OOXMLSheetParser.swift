import Foundation
import CoreGraphics

/// Parses `.xlsx` files (OOXML Spreadsheet format).
actor OOXMLSheetParser: DocReadable {
    nonisolated let url: URL
    nonisolated let format: DocumentFormat = .xlsx

    private let extractor: OOXMLZipExtractor

    private var _pageCount: Int?
    private var _pageSize: CGSize?
    private var _metadata: DocumentMetadata?

    init(url: URL) {
        self.url = url
        self.extractor = OOXMLZipExtractor(url: url)
    }

    var pageCount: Int {
        get throws {
            if let cached = _pageCount { return cached }
            try extractor.validateOOXMLStructure()
            let count = try parseSheetCount()
            _pageCount = count
            return count
        }
    }

    func pageSize(at index: Int) throws -> CGSize {
        if let cached = _pageSize { return cached }
        try extractor.validateOOXMLStructure()
        let size = try parsePageSize()
        _pageSize = size
        return size
    }

    var metadata: DocumentMetadata {
        get throws {
            if let cached = _metadata { return cached }
            let meta = try parseCoreProperties()
            _metadata = meta
            return meta
        }
    }

    func exportPDF() async throws -> Data {
        let count = try pageCount
        guard count > 0 else { throw DocReaderError.corruptedFile }
        return try await exportPDF(pages: 0...(count - 1))
    }

    func exportPDF(pages: ClosedRange<Int>) async throws -> Data {
        let count = try pageCount
        guard pages.lowerBound >= 0, pages.upperBound < count else {
            throw DocReaderError.pageOutOfRange
        }
        return try await PDFExporter.export(parser: self, pages: pages)
    }

    // MARK: - Parsing

    private func parseSheetCount() throws -> Int {
        let data = try extractor.extractEntry(path: "xl/workbook.xml")
        let handler = XLSXSheetCounter()
        let parser = XMLParser(data: data)
        parser.delegate = handler
        parser.parse()
        return max(1, handler.count)
    }

    private func parsePageSize() throws -> CGSize {
        // Try reading first worksheet for pageSetup
        let data = try extractor.extractEntry(path: "xl/worksheets/sheet1.xml")
        let handler = XLSXPageSetupParser()
        let parser = XMLParser(data: data)
        parser.delegate = handler
        parser.parse()
        // Map ECMA-376 paper size enum to points (A4 default)
        return handler.cgSize ?? CGSize(width: 595.28, height: 841.89)
    }

    private func parseCoreProperties() throws -> DocumentMetadata {
        guard let data = try? extractor.extractEntry(path: "docProps/core.xml") else {
            return DocumentMetadata()
        }
        let handler = XLSXCoreParser()
        let parser = XMLParser(data: data)
        parser.delegate = handler
        parser.parse()

        let formatter = ISO8601DateFormatter()
        return DocumentMetadata(
            title: handler.title,
            author: handler.creator,
            modifiedDate: handler.modified.flatMap { formatter.date(from: $0) },
            createdDate: handler.created.flatMap { formatter.date(from: $0) }
        )
    }
}

// MARK: - DocContentProviding

extension OOXMLSheetParser: DocContentProviding {
    func buildPageContents() throws -> [PageContent] {
        try extractor.validateOOXMLStructure()

        // Load shared strings (optional — not all xlsx files have this)
        var sharedStrings: [String] = []
        if let ssData = try? extractor.extractEntry(path: "xl/sharedStrings.xml") {
            let ssHandler = XLSXSharedStringsParser()
            let ssParser = XMLParser(data: ssData)
            ssParser.delegate = ssHandler
            ssParser.parse()
            sharedStrings = ssHandler.strings
        }

        // Load workbook relationships
        guard let relsData = try? extractor.extractEntry(path: "xl/_rels/workbook.xml.rels") else {
            return [.sheet(SheetPageContent(sheetName: "Sheet1", cells: []))]
        }
        let relsHandler = XLSXWorkbookRelsParser()
        let relsParser = XMLParser(data: relsData)
        relsParser.delegate = relsHandler
        relsParser.parse()

        // Load workbook (ordered sheets)
        let wbData = try extractor.extractEntry(path: "xl/workbook.xml")
        let wbHandler = XLSXWorkbookParser()
        let wbParser = XMLParser(data: wbData)
        wbParser.delegate = wbHandler
        wbParser.parse()

        var pages: [PageContent] = []
        for sheet in wbHandler.sheets {
            guard let path = relsHandler.relIdToPath[sheet.rId],
                  let sheetData = try? extractor.extractEntry(path: path) else { continue }

            let cellHandler = XLSXCellParser(sharedStrings: sharedStrings)
            let cellParser = XMLParser(data: sheetData)
            cellParser.delegate = cellHandler
            cellParser.parse()

            pages.append(.sheet(SheetPageContent(sheetName: sheet.name, cells: cellHandler.cells)))
        }

        return pages.isEmpty ? [.sheet(SheetPageContent(sheetName: "Sheet1", cells: []))] : pages
    }
}

// MARK: - Internal helpers (accessible for tests)

/// Converts a 0-based column index to an Excel-style column label (A, B, …, Z, AA, …).
func columnLabel(_ index: Int) -> String {
    var result = ""
    var n = index + 1
    while n > 0 {
        n -= 1
        result = String(UnicodeScalar(65 + (n % 26))!) + result
        n /= 26
    }
    return result
}

/// Extracts the 0-based column index from an Excel cell reference such as "B3" or "AA1".
func columnIndex(from ref: String) -> Int {
    var result = 0
    for char in ref.uppercased() {
        guard char.isLetter, let ascii = char.asciiValue else { break }
        result = result * 26 + Int(ascii - 64)
    }
    return result - 1
}

// MARK: - SAX Helpers

private final class XLSXSheetCounter: NSObject, XMLParserDelegate, @unchecked Sendable {
    var count = 0

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        if elementName == "sheet" {
            count += 1
        }
    }
}

/// Maps ECMA-376 `paperSize` attribute to CGSize in points.
private final class XLSXPageSetupParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    var cgSize: CGSize?

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        guard elementName == "pageSetup", cgSize == nil else { return }
        // ECMA-376 paper sizes (selected common values)
        let paperSize = Int(attributes["paperSize"] ?? "9") ?? 9
        cgSize = Self.pointSize(forECMAPaperSize: paperSize)
    }

    private static func pointSize(forECMAPaperSize code: Int) -> CGSize {
        // Source: ECMA-376 §18.18.43 ST_PaperSize
        let inch: CGFloat = 72.0
        switch code {
        case 1:  return CGSize(width: 8.5 * inch, height: 11 * inch)   // US Letter
        case 9:  return CGSize(width: 8.5 * inch, height: 11 * inch)   // US Letter (default)
        case 5:  return CGSize(width: 8.5 * inch, height: 14 * inch)   // US Legal
        case 8:  return CGSize(width: 8.27 * inch, height: 11.69 * inch) // A4
        case 26: return CGSize(width: 8.27 * inch, height: 11.69 * inch) // A4 small
        default: return CGSize(width: 595.28, height: 841.89)            // A4 fallback
        }
    }
}

private final class XLSXCoreParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    var title: String?
    var creator: String?
    var modified: String?
    var created: String?

    private var currentElement = ""
    private var currentValue = ""

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        currentElement = elementName
        currentValue = ""
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        currentValue += string
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName: String?) {
        let value = currentValue.trimmingCharacters(in: .whitespaces)
        switch elementName {
        case "dc:title":         title = value
        case "dc:creator":       creator = value
        case "dcterms:modified": modified = value
        case "dcterms:created":  created = value
        default: break
        }
    }
}

/// Parses xl/sharedStrings.xml to extract the shared string table.
private final class XLSXSharedStringsParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var strings: [String] = []

    private var currentSIText = ""
    private var currentText = ""
    private var inT = false

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        switch elementName {
        case "si":
            currentSIText = ""
        case "t":
            inT = true
            currentText = ""
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inT { currentText += string }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName: String?) {
        switch elementName {
        case "t":
            inT = false
            currentSIText += currentText
        case "si":
            strings.append(currentSIText)
        default:
            break
        }
    }
}

/// Parses xl/_rels/workbook.xml.rels to map relationship IDs to sheet paths.
private final class XLSXWorkbookRelsParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var relIdToPath: [String: String] = [:]

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        guard elementName == "Relationship" else { return }
        let type = attributes["Type"] ?? ""
        guard type.contains("worksheet"), !type.contains("chartsheet") else { return }
        guard let id = attributes["Id"], let target = attributes["Target"] else { return }

        // Target is relative to xl/ directory (e.g. "worksheets/sheet1.xml")
        let path: String
        if target.hasPrefix("xl/") || target.hasPrefix("/xl/") {
            path = target.hasPrefix("/") ? String(target.dropFirst()) : target
        } else if target.hasPrefix("../") {
            path = "xl/" + target.dropFirst(3)
        } else {
            path = "xl/\(target)"
        }
        relIdToPath[id] = path
    }
}

/// Parses xl/workbook.xml to get the ordered list of sheet names and relationship IDs.
private final class XLSXWorkbookParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var sheets: [(name: String, rId: String)] = []

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        guard elementName == "sheet" else { return }
        let name = attributes["name"] ?? "Sheet"
        let rId = attributes["r:id"] ?? attributes["rId"] ?? ""
        if !rId.isEmpty {
            sheets.append((name: name, rId: rId))
        }
    }
}

/// Parses a worksheet XML and extracts cell values.
private final class XLSXCellParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private let sharedStrings: [String]
    private(set) var cells: [SheetCellContent] = []

    private var currentRef = ""
    private var currentType = ""
    private var currentValue = ""
    private var inV = false

    init(sharedStrings: [String]) {
        self.sharedStrings = sharedStrings
    }

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        switch elementName {
        case "c":
            currentRef = attributes["r"] ?? ""
            currentType = attributes["t"] ?? ""
            currentValue = ""
        case "v":
            inV = true
        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inV { currentValue += string }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName: String?) {
        switch elementName {
        case "v":
            inV = false
        case "c":
            guard !currentRef.isEmpty else { return }
            let text: String
            switch currentType {
            case "s":
                let idx = Int(currentValue) ?? 0
                text = idx < sharedStrings.count ? sharedStrings[idx] : currentValue
            case "b":
                text = currentValue == "1" ? "TRUE" : "FALSE"
            default:
                text = currentValue
            }
            guard !text.isEmpty else { return }
            let col = columnIndex(from: currentRef)
            let rowStr = currentRef.drop(while: { $0.isLetter })
            let row = max(0, (Int(rowStr) ?? 1) - 1)
            cells.append(SheetCellContent(col: col, row: row, text: text))
        default:
            break
        }
    }
}
