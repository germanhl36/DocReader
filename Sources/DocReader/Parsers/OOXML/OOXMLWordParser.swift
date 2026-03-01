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
        // Fallback: count explicit <w:br w:type="page"/> breaks in word/document.xml
        let docData = try extractor.extractEntry(path: "word/document.xml")
        let handler = OOXMLPageBreakCounter()
        let parser = XMLParser(data: docData)
        parser.delegate = handler
        parser.parse()
        return handler.count + 1
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

// MARK: - DocContentProviding

extension OOXMLWordParser: DocContentProviding {
    func buildPageContents() throws -> [PageContent] {
        try extractor.validateOOXMLStructure()
        let pageSize = try parsePageSize()

        let docData = try extractor.extractEntry(path: "word/document.xml")
        let handler = OOXMLDocumentBodyParser()
        let xmlParser = XMLParser(data: docData)
        xmlParser.delegate = handler
        xmlParser.parse()

        let pages = splitIntoWordPages(
            paragraphs: handler.paragraphs,
            pageSize: pageSize,
            margins: handler.pageMargins
        )
        return pages.map { .word($0) }
    }
}

// MARK: - Internal helpers (accessible for tests)

/// Splits a flat paragraph array into pages at `__pagebreak__` sentinels.
func splitIntoWordPages(
    paragraphs: [WordParagraphContent],
    pageSize: CGSize,
    margins: WordPageMargins
) -> [WordPageContent] {
    var pages: [WordPageContent] = []
    var current: [WordParagraphContent] = []

    for para in paragraphs {
        if para.styleName == "__pagebreak__" {
            pages.append(WordPageContent(paragraphs: current, pageSize: pageSize, margins: margins))
            current = []
        } else {
            current.append(para)
        }
    }
    pages.append(WordPageContent(paragraphs: current, pageSize: pageSize, margins: margins))
    return pages
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

/// Counts explicit page breaks (`<w:br w:type="page"/>`) to determine page count.
private final class OOXMLPageBreakCounter: NSObject, XMLParserDelegate, @unchecked Sendable {
    var count = 0

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        if elementName == "w:br" || elementName == "br" {
            let breakType = attributes["w:type"] ?? attributes["type"]
            if breakType == "page" { count += 1 }
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

/// Full SAX parser for word/document.xml — extracts paragraphs with run-level formatting.
private final class OOXMLDocumentBodyParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var paragraphs: [WordParagraphContent] = []
    private(set) var pageMargins = WordPageMargins(top: 72, bottom: 72, left: 72, right: 72)

    // Paragraph state
    private var inParagraph = false
    private var currentParaStyle = "Normal"
    private var currentParaSpacingAfter: CGFloat = 0
    private var currentRuns: [WordRunContent] = []

    // Run state
    private var inRun = false
    private var currentRunBold = false
    private var currentRunItalic = false
    private var currentRunFontSize: CGFloat = 12
    private var currentRunColor: String? = nil
    private var currentRunText = ""
    private var inText = false

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        switch elementName {
        case "w:p", "p":
            inParagraph = true
            currentParaStyle = "Normal"
            currentParaSpacingAfter = 0
            currentRuns = []

        case "w:r", "r":
            guard inParagraph else { return }
            inRun = true
            currentRunBold = false
            currentRunItalic = false
            currentRunFontSize = 12
            currentRunColor = nil
            currentRunText = ""

        case "w:pStyle", "pStyle":
            let val = attributes["w:val"] ?? attributes["val"] ?? "Normal"
            currentParaStyle = val

        case "w:spacing", "spacing":
            if let afterStr = attributes["w:after"] ?? attributes["after"],
               let afterTwips = Int(afterStr) {
                currentParaSpacingAfter = CGFloat(afterTwips) / 20.0
            }

        case "w:b", "b":
            guard inRun else { return }
            let val = attributes["w:val"] ?? attributes["val"]
            currentRunBold = val != "0"

        case "w:i", "i":
            guard inRun else { return }
            let val = attributes["w:val"] ?? attributes["val"]
            currentRunItalic = val != "0"

        case "w:sz", "sz":
            guard inRun else { return }
            if let szStr = attributes["w:val"] ?? attributes["val"],
               let sz = Int(szStr) {
                currentRunFontSize = CGFloat(sz) / 2.0
            }

        case "w:color", "color":
            guard inRun else { return }
            let val = attributes["w:val"] ?? attributes["val"]
            if val != "auto" { currentRunColor = val }

        case "w:t", "t":
            if inRun { inText = true }

        case "w:br", "br":
            let breakType = attributes["w:type"] ?? attributes["type"]
            guard breakType == "page", inParagraph else { return }
            // Flush any pending runs into a paragraph before the break
            if !currentRuns.isEmpty {
                paragraphs.append(WordParagraphContent(
                    runs: currentRuns,
                    styleName: currentParaStyle,
                    spacingAfterPt: currentParaSpacingAfter
                ))
                currentRuns = []
            }
            // Sentinel marks a page boundary
            paragraphs.append(WordParagraphContent(
                runs: [],
                styleName: "__pagebreak__",
                spacingAfterPt: 0
            ))

        case "w:pgMar", "pgMar":
            let topTwips    = Int(attributes["w:top"]    ?? attributes["top"]    ?? "1440") ?? 1440
            let bottomTwips = Int(attributes["w:bottom"] ?? attributes["bottom"] ?? "1440") ?? 1440
            let leftTwips   = Int(attributes["w:left"]   ?? attributes["left"]   ?? "1440") ?? 1440
            let rightTwips  = Int(attributes["w:right"]  ?? attributes["right"]  ?? "1440") ?? 1440
            pageMargins = WordPageMargins(
                top:    CGFloat(topTwips)    / 20.0,
                bottom: CGFloat(bottomTwips) / 20.0,
                left:   CGFloat(leftTwips)   / 20.0,
                right:  CGFloat(rightTwips)  / 20.0
            )

        default:
            break
        }
    }

    func parser(_ parser: XMLParser, foundCharacters string: String) {
        if inText { currentRunText += string }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName: String?) {
        switch elementName {
        case "w:t", "t":
            inText = false

        case "w:r", "r":
            guard inRun else { return }
            if !currentRunText.isEmpty {
                currentRuns.append(WordRunContent(
                    text: currentRunText,
                    bold: currentRunBold,
                    italic: currentRunItalic,
                    fontSizePt: currentRunFontSize,
                    hexColor: currentRunColor
                ))
            }
            inRun = false

        case "w:p", "p":
            guard inParagraph else { return }
            if !currentRuns.isEmpty {
                paragraphs.append(WordParagraphContent(
                    runs: currentRuns,
                    styleName: currentParaStyle,
                    spacingAfterPt: currentParaSpacingAfter
                ))
            }
            inParagraph = false

        default:
            break
        }
    }
}
