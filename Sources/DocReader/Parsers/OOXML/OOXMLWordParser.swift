import Foundation
import CoreGraphics
import CoreText

/// Parses `.docx` files (OOXML Word format).
actor OOXMLWordParser: DocReadable {
    nonisolated let url: URL
    nonisolated let format: DocumentFormat = .docx

    private let extractor: OOXMLZipExtractor

    // Cached values
    private var _pageCount: Int?
    private var _pageSize: CGSize?
    private var _metadata: DocumentMetadata?
    private var _pageContents: [PageContent]?

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
        // Build content first to get the actual reflowed page count, which may exceed the
        // value recorded in docProps/app.xml when overflow reflow creates additional pages.
        let contents = try buildPageContents()
        guard !contents.isEmpty else { throw DocReaderError.corruptedFile }
        let pageSize = try parsePageSize()
        return try await PDFExporter.exportContents(contents, pageSize: pageSize)
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
        if let cached = _pageContents { return cached }
        try extractor.validateOOXMLStructure()
        let pageSize = try parsePageSize()

        // Parse numbering.xml (optional — not all documents have lists)
        let numberingParser = OOXMLNumberingParser()
        if let numData = try? extractor.extractEntry(path: "word/numbering.xml") {
            let numXml = XMLParser(data: numData)
            numXml.delegate = numberingParser
            numXml.parse()
        }

        let docData = try extractor.extractEntry(path: "word/document.xml")
        let handler = OOXMLDocumentBodyParser(numbering: numberingParser)
        let xmlParser = XMLParser(data: docData)
        xmlParser.delegate = handler
        xmlParser.parse()

        let pages = splitIntoWordPages(
            elements: handler.elements,
            pageSize: pageSize,
            margins: handler.pageMargins
        )
        let result = pages.map { PageContent.word($0) }
        _pageContents = result
        return result
    }
}

// MARK: - Internal helpers (accessible for tests)

/// Splits a flat element array into pages, honouring explicit `__pagebreak__` sentinels and
/// automatically starting a new page when content overflows the available vertical space.
func splitIntoWordPages(
    elements: [WordElement],
    pageSize: CGSize,
    margins: WordPageMargins
) -> [WordPageContent] {
    let contentHeight = pageSize.height - margins.top - margins.bottom
    let contentWidth  = pageSize.width  - margins.left - margins.right

    var pages: [WordPageContent] = []
    var current: [WordElement] = []
    var remainingY: CGFloat = contentHeight

    for element in elements {
        // Explicit page break — flush and start a new page
        if case .paragraph(let para) = element, para.styleName == "__pagebreak__" {
            pages.append(WordPageContent(elements: current, pageSize: pageSize, margins: margins))
            current = []
            remainingY = contentHeight
            continue
        }

        let elementHeight = WordPageRenderer.measureElement(element, availableWidth: contentWidth)

        // Overflow — start a new page before adding this element
        if !current.isEmpty && elementHeight > 0 && remainingY - elementHeight < 0 {
            pages.append(WordPageContent(elements: current, pageSize: pageSize, margins: margins))
            current = []
            remainingY = contentHeight
        }

        current.append(element)
        remainingY -= elementHeight
    }

    pages.append(WordPageContent(elements: current, pageSize: pageSize, margins: margins))
    return pages
}

// MARK: - Numbering parser (word/numbering.xml)

final class OOXMLNumberingParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    // abstractNumId -> [ilvl -> (format, startVal)]
    private var abstractNums: [Int: [Int: (format: String, startVal: Int)]] = [:]
    // numId -> abstractNumId
    private var nums: [Int: Int] = [:]

    private var currentAbstractNumId: Int?
    private var currentIlvl: Int?
    private var currentFormat: String?
    private var currentStartVal: Int?
    private var currentNumId: Int?
    private var currentAbstractNumIdRef: Int?

    func format(numId: Int, ilvl: Int) -> String? {
        guard let abstractId = nums[numId],
              let levels = abstractNums[abstractId],
              let level = levels[ilvl] else { return nil }
        return level.format
    }

    func startVal(numId: Int, ilvl: Int) -> Int {
        guard let abstractId = nums[numId],
              let levels = abstractNums[abstractId],
              let level = levels[ilvl] else { return 1 }
        return level.startVal
    }

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        switch elementName {
        case "w:abstractNum", "abstractNum":
            let idStr = attributes["w:abstractNumId"] ?? attributes["abstractNumId"] ?? ""
            currentAbstractNumId = Int(idStr)
            currentIlvl = nil

        case "w:lvl", "lvl":
            let idStr = attributes["w:ilvl"] ?? attributes["ilvl"] ?? ""
            currentIlvl = Int(idStr)
            currentFormat = nil
            currentStartVal = nil

        case "w:numFmt", "numFmt":
            currentFormat = attributes["w:val"] ?? attributes["val"]

        case "w:start", "start":
            if let v = Int(attributes["w:val"] ?? attributes["val"] ?? "") {
                currentStartVal = v
            }

        case "w:num", "num":
            let idStr = attributes["w:numId"] ?? attributes["numId"] ?? ""
            currentNumId = Int(idStr)
            currentAbstractNumIdRef = nil

        case "w:abstractNumId", "abstractNumId":
            if let v = Int(attributes["w:val"] ?? attributes["val"] ?? "") {
                currentAbstractNumIdRef = v
            }

        default:
            break
        }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName: String?) {
        switch elementName {
        case "w:lvl", "lvl":
            if let abstractId = currentAbstractNumId, let ilvl = currentIlvl,
               let fmt = currentFormat {
                if abstractNums[abstractId] == nil { abstractNums[abstractId] = [:] }
                abstractNums[abstractId]![ilvl] = (format: fmt, startVal: currentStartVal ?? 1)
            }
            currentIlvl = nil
            currentFormat = nil
            currentStartVal = nil

        case "w:abstractNum", "abstractNum":
            currentAbstractNumId = nil

        case "w:num", "num":
            if let numId = currentNumId, let abstractId = currentAbstractNumIdRef {
                nums[numId] = abstractId
            }
            currentNumId = nil
            currentAbstractNumIdRef = nil

        default:
            break
        }
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

/// Full SAX parser for word/document.xml — extracts elements (paragraphs + tables) with formatting.
final class OOXMLDocumentBodyParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var elements: [WordElement] = []
    private(set) var pageMargins = WordPageMargins(top: 72, bottom: 72, left: 72, right: 72)

    private let numbering: OOXMLNumberingParser
    // List counters: key = "numId-ilvl"
    private var listCounters: [String: Int] = [:]

    init(numbering: OOXMLNumberingParser) {
        self.numbering = numbering
    }

    // Paragraph state
    private var inParagraph = false
    private var currentParaStyle = "Normal"
    private var currentParaSpacingAfter: CGFloat = 0
    private var currentParaSpacingBefore: CGFloat = 0
    private var currentParaAlignment: CTTextAlignment = .natural
    private var currentParaLeftIndent: CGFloat = 0
    private var currentParaFirstLineIndent: CGFloat = 0
    private var currentParaBackground: String? = nil
    private var currentRuns: [WordRunContent] = []

    // List state (inside w:numPr)
    private var inNumPr = false
    private var currentListNumId: Int? = nil
    private var currentListIlvl: Int = 0

    // Run state
    private var inRun = false
    private var currentRunBold = false
    private var currentRunItalic = false
    private var currentRunFontSize: CGFloat = 12
    private var currentRunColor: String? = nil
    private var currentRunText = ""
    private var inText = false

    // Table state
    private var inTable = false
    private var inRow = false
    private var inCell = false
    private var currentTableRows: [WordTableRow] = []
    private var currentRowCells: [WordTableCell] = []
    private var currentCellParagraphs: [WordParagraphContent] = []

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        switch elementName {

        // Table structure
        case "w:tbl", "tbl":
            inTable = true
            currentTableRows = []

        case "w:tr", "tr":
            guard inTable else { return }
            inRow = true
            currentRowCells = []

        case "w:tc", "tc":
            guard inTable, inRow else { return }
            inCell = true
            currentCellParagraphs = []

        // Paragraph
        case "w:p", "p":
            inParagraph = true
            currentParaStyle = "Normal"
            currentParaSpacingAfter = 0
            currentParaSpacingBefore = 0
            currentParaAlignment = .natural
            currentParaLeftIndent = 0
            currentParaFirstLineIndent = 0
            currentParaBackground = nil
            currentRuns = []
            currentListNumId = nil
            currentListIlvl = 0

        case "w:r", "r":
            guard inParagraph else { return }
            inRun = true
            currentRunBold = false
            currentRunItalic = false
            currentRunFontSize = 12
            currentRunColor = nil
            currentRunText = ""

        // Paragraph properties
        case "w:pStyle", "pStyle":
            let val = attributes["w:val"] ?? attributes["val"] ?? "Normal"
            currentParaStyle = val

        case "w:jc", "jc":
            let val = attributes["w:val"] ?? attributes["val"] ?? ""
            switch val {
            case "left":       currentParaAlignment = .left
            case "right":      currentParaAlignment = .right
            case "center":     currentParaAlignment = .center
            case "both", "distribute": currentParaAlignment = .justified
            default:           currentParaAlignment = .natural
            }

        case "w:spacing", "spacing":
            if let afterStr = attributes["w:after"] ?? attributes["after"],
               let afterTwips = Int(afterStr) {
                currentParaSpacingAfter = CGFloat(afterTwips) / 20.0
            }
            if let beforeStr = attributes["w:before"] ?? attributes["before"],
               let beforeTwips = Int(beforeStr) {
                currentParaSpacingBefore = CGFloat(beforeTwips) / 20.0
            }

        case "w:ind", "ind":
            if let leftStr = attributes["w:left"] ?? attributes["left"],
               let leftTwips = Int(leftStr) {
                currentParaLeftIndent = CGFloat(leftTwips) / 20.0
            }
            if let firstStr = attributes["w:firstLine"] ?? attributes["firstLine"],
               let firstTwips = Int(firstStr) {
                currentParaFirstLineIndent = CGFloat(firstTwips) / 20.0
            }

        case "w:shd", "shd":
            let val  = attributes["w:val"]   ?? attributes["val"]   ?? ""
            let fill = attributes["w:fill"]  ?? attributes["fill"]  ?? ""
            let shdColor = attributes["w:color"] ?? attributes["color"] ?? ""
            // "solid" pattern: the foreground (w:color) covers the whole area.
            // All other patterns (clear, pct*, etc.): use the background fill (w:fill).
            let bgHex = (val == "solid") ? shdColor : fill
            if !bgHex.isEmpty && bgHex != "auto" && bgHex.uppercased() != "FFFFFF" {
                currentParaBackground = bgHex
            }

        case "w:numPr", "numPr":
            inNumPr = true

        case "w:ilvl", "ilvl":
            guard inNumPr else { return }
            if let v = Int(attributes["w:val"] ?? attributes["val"] ?? "") {
                currentListIlvl = v
            }

        case "w:numId", "numId":
            guard inNumPr else { return }
            if let v = Int(attributes["w:val"] ?? attributes["val"] ?? "") {
                currentListNumId = v
            }

        // Run properties
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
            guard breakType == "page", inParagraph, !inCell else { return }
            // Flush pending runs before inserting page break sentinel
            if !currentRuns.isEmpty {
                elements.append(.paragraph(makeParagraph()))
                currentRuns = []
            }
            elements.append(.paragraph(WordParagraphContent(
                runs: [],
                styleName: "__pagebreak__",
                spacingAfterPt: 0
            )))

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

        case "w:tbl", "tbl":
            guard inTable else { return }
            elements.append(.table(WordTableContent(rows: currentTableRows)))
            currentTableRows = []
            inTable = false

        case "w:tr", "tr":
            guard inTable, inRow else { return }
            let isHeader = currentTableRows.isEmpty
            currentTableRows.append(WordTableRow(cells: currentRowCells, isHeader: isHeader))
            currentRowCells = []
            inRow = false

        case "w:tc", "tc":
            guard inTable, inCell else { return }
            currentRowCells.append(WordTableCell(paragraphs: currentCellParagraphs))
            currentCellParagraphs = []
            inCell = false

        case "w:numPr", "numPr":
            inNumPr = false

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
                let para = makeParagraph()
                if inCell {
                    currentCellParagraphs.append(para)
                } else {
                    elements.append(.paragraph(para))
                }
            }
            inParagraph = false

        default:
            break
        }
    }

    // MARK: - Private helpers

    private func makeParagraph() -> WordParagraphContent {
        let prefix = resolveListPrefix()
        return WordParagraphContent(
            runs: currentRuns,
            styleName: currentParaStyle,
            spacingAfterPt: currentParaSpacingAfter,
            alignment: currentParaAlignment,
            listPrefix: prefix,
            spacingBeforePt: currentParaSpacingBefore,
            leftIndentPt: currentParaLeftIndent,
            firstLineIndentPt: currentParaFirstLineIndent,
            backgroundHex: currentParaBackground
        )
    }

    private func resolveListPrefix() -> String? {
        guard let numId = currentListNumId else { return nil }
        let ilvl = currentListIlvl
        guard let format = numbering.format(numId: numId, ilvl: ilvl) else { return nil }

        let key = "\(numId)-\(ilvl)"
        let counter = (listCounters[key] ?? (numbering.startVal(numId: numId, ilvl: ilvl) - 1)) + 1
        listCounters[key] = counter

        switch format {
        case "bullet":
            return ilvl == 0 ? "• " : "◦ "
        case "decimal":
            return "\(counter). "
        case "lowerLetter":
            return "\(letterLabel(counter - 1)). "
        case "upperLetter":
            return "\(letterLabel(counter - 1).uppercased()). "
        case "lowerRoman":
            return "\(romanNumeral(counter).lowercased()). "
        case "upperRoman":
            return "\(romanNumeral(counter)). "
        default:
            return "• "
        }
    }

    private func letterLabel(_ index: Int) -> String {
        let letters = Array("abcdefghijklmnopqrstuvwxyz")
        guard index >= 0 else { return "a" }
        return String(letters[index % 26])
    }

    private func romanNumeral(_ n: Int) -> String {
        let values = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        let symbols = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"]
        var result = ""
        var remaining = n
        for (val, sym) in zip(values, symbols) {
            while remaining >= val {
                result += sym
                remaining -= val
            }
        }
        return result.isEmpty ? "I" : result
    }
}
