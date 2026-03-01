import Foundation
import CoreGraphics

/// Parses `.pptx` files (OOXML Presentation format).
actor OOXMLSlideParser: DocReadable {
    nonisolated let url: URL
    nonisolated let format: DocumentFormat = .pptx

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
            let count = try parseSlideCount()
            _pageCount = count
            return count
        }
    }

    func pageSize(at index: Int) throws -> CGSize {
        if let cached = _pageSize { return cached }
        try extractor.validateOOXMLStructure()
        let size = try parseSlideSize()
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

    private func parseSlideCount() throws -> Int {
        let data = try extractor.extractEntry(path: "ppt/presentation.xml")
        let handler = PPTXSlideIdCounter()
        let parser = XMLParser(data: data)
        parser.delegate = handler
        parser.parse()
        return max(1, handler.count)
    }

    private func parseSlideSize() throws -> CGSize {
        let data = try extractor.extractEntry(path: "ppt/presentation.xml")
        let handler = PPTXSlideSizeParser()
        let parser = XMLParser(data: data)
        parser.delegate = handler
        parser.parse()
        if let cx = handler.cxEMU, let cy = handler.cyEMU, cx > 0, cy > 0 {
            // EMU to points: divide by 12700 (1 pt = 12700 EMU)
            return CGSize(width: CGFloat(cx) / 12700.0, height: CGFloat(cy) / 12700.0)
        }
        // Default: widescreen 16:9
        return CGSize(width: 720, height: 405)
    }

    private func parseCoreProperties() throws -> DocumentMetadata {
        guard let data = try? extractor.extractEntry(path: "docProps/core.xml") else {
            return DocumentMetadata()
        }
        let handler = PPTXCoreParser()
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

extension OOXMLSlideParser: DocContentProviding {
    func buildPageContents() throws -> [PageContent] {
        try extractor.validateOOXMLStructure()

        // Load slide relationship IDs → paths
        guard let relsData = try? extractor.extractEntry(path: "ppt/_rels/presentation.xml.rels") else {
            return [.slide(SlidePageContent(textBoxes: []))]
        }
        let relsHandler = PPTXRelationshipParser()
        let relsParser = XMLParser(data: relsData)
        relsParser.delegate = relsHandler
        relsParser.parse()

        // Load presentation to get ordered slide rIds
        let presData = try extractor.extractEntry(path: "ppt/presentation.xml")
        let orderHandler = PPTXSlideOrderParser()
        let orderParser = XMLParser(data: presData)
        orderParser.delegate = orderHandler
        orderParser.parse()

        var pages: [PageContent] = []
        for rId in orderHandler.orderedRIds {
            guard let path = relsHandler.slideRelIds[rId],
                  let slideData = try? extractor.extractEntry(path: path) else { continue }

            let contentHandler = PPTXSlideContentParser()
            let contentParser = XMLParser(data: slideData)
            contentParser.delegate = contentHandler
            contentParser.parse()

            pages.append(.slide(SlidePageContent(textBoxes: contentHandler.textBoxes)))
        }

        return pages.isEmpty ? [.slide(SlidePageContent(textBoxes: []))] : pages
    }
}

// MARK: - SAX Helpers

private final class PPTXSlideIdCounter: NSObject, XMLParserDelegate, @unchecked Sendable {
    var count = 0

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        if elementName == "p:sldId" || elementName == "sldId" {
            count += 1
        }
    }
}

private final class PPTXSlideSizeParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    var cxEMU: Int?
    var cyEMU: Int?

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        if elementName == "p:sldSz" || elementName == "sldSz" {
            if let cx = attributes["cx"] { cxEMU = Int(cx) }
            if let cy = attributes["cy"] { cyEMU = Int(cy) }
        }
    }
}

private final class PPTXCoreParser: NSObject, XMLParserDelegate, @unchecked Sendable {
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

/// Parses ppt/_rels/presentation.xml.rels to map slide rIds to their paths.
private final class PPTXRelationshipParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var slideRelIds: [String: String] = [:]

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        guard elementName == "Relationship" else { return }
        let type = attributes["Type"] ?? ""
        // Match exactly the slide relationship type (not slideLayout, slideMaster, etc.)
        guard type.hasSuffix("/slide") else { return }
        guard let id = attributes["Id"], let target = attributes["Target"] else { return }

        // Normalize target relative to ppt/ directory
        let path: String
        if target.hasPrefix("../") {
            path = "ppt/" + target.dropFirst(3)
        } else if target.hasPrefix("/") {
            path = String(target.dropFirst())
        } else {
            path = "ppt/\(target)"
        }
        slideRelIds[id] = path
    }
}

/// Reads the ordered r:id values from <p:sldId> in ppt/presentation.xml.
private final class PPTXSlideOrderParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var orderedRIds: [String] = []

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        if elementName == "p:sldId" || elementName == "sldId" {
            let rId = attributes["r:id"] ?? attributes["rId"] ?? ""
            if !rId.isEmpty { orderedRIds.append(rId) }
        }
    }
}

/// Parses a slide XML and extracts positioned text boxes.
private final class PPTXSlideContentParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var textBoxes: [SlideTextBoxContent] = []

    private var inSP = false
    private var isTitle = false
    private var offX = 0
    private var offY = 0
    private var extCX = 0
    private var extCY = 0
    private var lines: [String] = []
    private var inTxBody = false
    private var inParagraph = false
    private var currentParaText = ""
    private var inT = false
    private var currentText = ""

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        let local = elementName.components(separatedBy: ":").last ?? elementName

        switch local {
        case "sp":
            inSP = true
            isTitle = false
            offX = 0; offY = 0; extCX = 0; extCY = 0
            lines = []
            inTxBody = false

        case "ph":
            guard inSP else { return }
            let t = attributes["type"] ?? ""
            isTitle = (t == "title" || t == "ctrTitle")

        case "off":
            guard inSP else { return }
            offX = Int(attributes["x"] ?? "0") ?? 0
            offY = Int(attributes["y"] ?? "0") ?? 0

        case "ext":
            guard inSP else { return }
            extCX = Int(attributes["cx"] ?? "0") ?? 0
            extCY = Int(attributes["cy"] ?? "0") ?? 0

        case "txBody":
            if inSP { inTxBody = true }

        case "p":
            guard inTxBody else { return }
            inParagraph = true
            currentParaText = ""

        case "t":
            guard inTxBody else { return }
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
        let local = elementName.components(separatedBy: ":").last ?? elementName

        switch local {
        case "t":
            inT = false
            currentParaText += currentText

        case "p":
            if inParagraph, !currentParaText.isEmpty {
                lines.append(currentParaText)
            }
            inParagraph = false

        case "txBody":
            inTxBody = false

        case "sp":
            if inSP, !lines.isEmpty {
                let emuPt: CGFloat = 1.0 / 12700.0
                let frame = CGRect(
                    x: CGFloat(offX) * emuPt,
                    y: CGFloat(offY) * emuPt,
                    width:  CGFloat(extCX) * emuPt,
                    height: CGFloat(extCY) * emuPt
                )
                textBoxes.append(SlideTextBoxContent(frame: frame, lines: lines, isTitle: isTitle))
            }
            inSP = false

        default:
            break
        }
    }
}
