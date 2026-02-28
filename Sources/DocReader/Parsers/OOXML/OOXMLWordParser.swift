import Foundation
import CoreGraphics

/// Parses `.docx` files (OOXML Word format).
actor OOXMLWordParser: DocReadable {
    nonisolated let url: URL
    nonisolated let format: DocumentFormat = .docx

    private let extractor: OOXMLZipExtractor

    // Cached values
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
            let count = try parsePageCount()
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

    private func parsePageCount() throws -> Int {
        // Primary: docProps/app.xml <Pages>
        if let appData = try? extractor.extractEntry(path: "docProps/app.xml") {
            let handler = OOXMLAppPropertiesParser()
            let parser = XMLParser(data: appData)
            parser.delegate = handler
            parser.parse()
            if let pages = handler.pages, pages > 0 {
                return pages
            }
        }
        // Fallback: count <w:sectPr> in word/document.xml
        let docData = try extractor.extractEntry(path: "word/document.xml")
        let handler = OOXMLSectPrCounter()
        let parser = XMLParser(data: docData)
        parser.delegate = handler
        parser.parse()
        return max(1, handler.count)
    }

    private func parsePageSize() throws -> CGSize {
        let data = try extractor.extractEntry(path: "word/document.xml")
        let handler = OOXMLPageSizeParser()
        let parser = XMLParser(data: data)
        parser.delegate = handler
        parser.parse()
        if let width = handler.widthTwips, let height = handler.heightTwips, width > 0, height > 0 {
            // Twips to points: divide by 20
            return CGSize(width: CGFloat(width) / 20.0, height: CGFloat(height) / 20.0)
        }
        // Default to US Letter
        return CGSize(width: 612, height: 792)
    }

    private func parseCoreProperties() throws -> DocumentMetadata {
        guard let data = try? extractor.extractEntry(path: "docProps/core.xml") else {
            return DocumentMetadata()
        }
        let handler = OOXMLCorePropertiesParser()
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

// MARK: - SAX Helpers (private)

private final class OOXMLAppPropertiesParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    var pages: Int?
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
        if elementName == "Pages" {
            pages = Int(currentValue.trimmingCharacters(in: .whitespaces))
        }
    }
}

private final class OOXMLSectPrCounter: NSObject, XMLParserDelegate, @unchecked Sendable {
    var count = 0

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        // Local name may come with namespace prefix stripped
        if elementName == "w:sectPr" || elementName == "sectPr" {
            count += 1
        }
    }
}

private final class OOXMLPageSizeParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    var widthTwips: Int?
    var heightTwips: Int?

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        if elementName == "w:pgSz" || elementName == "pgSz" {
            if let w = attributes["w:w"] ?? attributes["w"] {
                widthTwips = Int(w)
            }
            if let h = attributes["w:h"] ?? attributes["h"] {
                heightTwips = Int(h)
            }
        }
    }
}

private final class OOXMLCorePropertiesParser: NSObject, XMLParserDelegate, @unchecked Sendable {
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
        case "dc:title":   title = value
        case "dc:creator": creator = value
        case "dcterms:modified": modified = value
        case "dcterms:created":  created = value
        default: break
        }
    }
}
