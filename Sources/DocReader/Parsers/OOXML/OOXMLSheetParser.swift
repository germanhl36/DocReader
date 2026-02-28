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
