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

        // Parse styles.xml (optional — provides default font/size/heading styles)
        let stylesParser = OOXMLStylesParser()
        if let stylesData = try? extractor.extractEntry(path: "word/styles.xml") {
            let stylesXml = XMLParser(data: stylesData)
            stylesXml.delegate = stylesParser
            stylesXml.parse()
        }

        // Parse numbering.xml (optional — not all documents have lists)
        let numberingParser = OOXMLNumberingParser()
        if let numData = try? extractor.extractEntry(path: "word/numbering.xml") {
            let numXml = XMLParser(data: numData)
            numXml.delegate = numberingParser
            numXml.parse()
        }

        let docData = try extractor.extractEntry(path: "word/document.xml")
        let handler = OOXMLDocumentBodyParser(numbering: numberingParser, styles: stylesParser)
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
    var prevSpacingAfterPt: CGFloat = 0

    for element in elements {
        // Explicit page break — flush and start a new page
        if case .paragraph(let para) = element, para.styleName == "__pagebreak__" {
            pages.append(WordPageContent(elements: current, pageSize: pageSize, margins: margins))
            current = []
            remainingY = contentHeight
            prevSpacingAfterPt = 0
            continue
        }

        var elementHeight = WordPageRenderer.measureElement(element, availableWidth: contentWidth)

        // Paragraph spacing collapse (Word rule): the effective gap between two adjacent
        // paragraphs is max(prevAfter, thisBefore), not the sum. Subtract the redundant
        // portion from this element's measured height for accurate page-break decisions.
        if case .paragraph(let para) = element {
            let collapsedSaving = min(para.spacingBeforePt, prevSpacingAfterPt)
            elementHeight -= collapsedSaving
            prevSpacingAfterPt = para.spacingAfterPt
        } else {
            prevSpacingAfterPt = 0
        }

        // Overflow — start a new page before adding this element
        if !current.isEmpty && elementHeight > 0 && remainingY - elementHeight < 0 {
            pages.append(WordPageContent(elements: current, pageSize: pageSize, margins: margins))
            current = []
            remainingY = contentHeight
            prevSpacingAfterPt = 0
        }

        current.append(element)
        remainingY -= elementHeight
    }

    pages.append(WordPageContent(elements: current, pageSize: pageSize, margins: margins))
    return pages
}

// MARK: - Numbering parser (word/numbering.xml)

private struct NumberingLevel {
    var format: String
    var startVal: Int
    var lvlText: String?
    var leftIndentPt: CGFloat?
    var hangingIndentPt: CGFloat?
}

final class OOXMLNumberingParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    // abstractNumId -> [ilvl -> NumberingLevel]
    private var abstractNums: [Int: [Int: NumberingLevel]] = [:]
    // numId -> abstractNumId
    private var nums: [Int: Int] = [:]

    private var currentAbstractNumId: Int?
    private var currentIlvl: Int?
    private var currentFormat: String?
    private var currentStartVal: Int?
    private var currentLvlText: String?
    private var currentLeftIndentPt: CGFloat?
    private var currentHangingIndentPt: CGFloat?
    private var currentNumId: Int?
    private var currentAbstractNumIdRef: Int?
    private var inLvl = false
    private var inLvlPPr = false

    func format(numId: Int, ilvl: Int) -> String? {
        level(numId: numId, ilvl: ilvl)?.format
    }

    func startVal(numId: Int, ilvl: Int) -> Int {
        level(numId: numId, ilvl: ilvl)?.startVal ?? 1
    }

    func lvlText(numId: Int, ilvl: Int) -> String? {
        level(numId: numId, ilvl: ilvl)?.lvlText
    }

    func leftIndent(numId: Int, ilvl: Int) -> CGFloat? {
        level(numId: numId, ilvl: ilvl)?.leftIndentPt
    }

    func hangingIndent(numId: Int, ilvl: Int) -> CGFloat? {
        level(numId: numId, ilvl: ilvl)?.hangingIndentPt
    }

    private func level(numId: Int, ilvl: Int) -> NumberingLevel? {
        guard let abstractId = nums[numId],
              let levels = abstractNums[abstractId] else { return nil }
        return levels[ilvl]
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
            currentLvlText = nil
            currentLeftIndentPt = nil
            currentHangingIndentPt = nil
            inLvl = true

        case "w:numFmt", "numFmt":
            currentFormat = attributes["w:val"] ?? attributes["val"]

        case "w:start", "start":
            if let v = Int(attributes["w:val"] ?? attributes["val"] ?? "") {
                currentStartVal = v
            }

        case "w:lvlText", "lvlText":
            guard inLvl else { break }
            currentLvlText = attributes["w:val"] ?? attributes["val"]

        case "w:pPr", "pPr":
            if inLvl { inLvlPPr = true }

        case "w:ind", "ind":
            guard inLvlPPr else { break }
            if let leftStr = attributes["w:left"] ?? attributes["left"],
               let leftTwips = Int(leftStr) {
                currentLeftIndentPt = CGFloat(leftTwips) / 20.0
            }
            if let hangStr = attributes["w:hanging"] ?? attributes["hanging"],
               let hangTwips = Int(hangStr) {
                currentHangingIndentPt = CGFloat(hangTwips) / 20.0
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
                abstractNums[abstractId]![ilvl] = NumberingLevel(
                    format: fmt,
                    startVal: currentStartVal ?? 1,
                    lvlText: currentLvlText,
                    leftIndentPt: currentLeftIndentPt,
                    hangingIndentPt: currentHangingIndentPt
                )
            }
            currentIlvl = nil
            currentFormat = nil
            currentStartVal = nil
            currentLvlText = nil
            currentLeftIndentPt = nil
            currentHangingIndentPt = nil
            inLvl = false

        case "w:pPr", "pPr":
            inLvlPPr = false

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

// MARK: - Styles parser (word/styles.xml)

struct WordStyleDefaults {
    var fontFamily: String? = nil
    var fontSizePt: CGFloat? = nil
    var bold: Bool = false
    var italic: Bool = false
    var color: String? = nil
    var spacingBeforePt: CGFloat? = nil
    var spacingAfterPt: CGFloat? = nil
}

final class OOXMLStylesParser: NSObject, XMLParserDelegate, @unchecked Sendable {
    private(set) var defaultFontFamily = "Arial"
    private(set) var defaultFontSizePt: CGFloat = 10
    private var styles: [String: WordStyleDefaults] = [:]

    private var inDocDefaults = false
    private var inRPrDefault = false
    private var currentStyleId: String?
    private var inStyleRPr = false
    private var inStylePPr = false
    private var current = WordStyleDefaults()
    private var currentSpacingBefore: CGFloat?
    private var currentSpacingAfter: CGFloat?

    func resolve(styleId: String) -> WordStyleDefaults? { styles[styleId] }

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        switch elementName {
        case "w:docDefaults", "docDefaults":
            inDocDefaults = true
        case "w:rPrDefault", "rPrDefault":
            if inDocDefaults { inRPrDefault = true }
        case "w:style", "style":
            currentStyleId = attributes["w:styleId"] ?? attributes["styleId"]
            current = WordStyleDefaults()
            currentSpacingBefore = nil; currentSpacingAfter = nil
            inStyleRPr = false; inStylePPr = false
        case "w:pPr", "pPr":
            if currentStyleId != nil { inStylePPr = true }
        case "w:rPr", "rPr":
            if currentStyleId != nil { inStyleRPr = true }
        case "w:rFonts", "rFonts":
            let family = attributes["w:ascii"] ?? attributes["ascii"]
                      ?? attributes["w:hAnsi"] ?? attributes["hAnsi"]
            guard let f = family, !f.isEmpty else { break }
            if inDocDefaults && inRPrDefault { defaultFontFamily = f }
            else if inStyleRPr { current.fontFamily = f }
        case "w:sz", "sz":
            guard let v = Int(attributes["w:val"] ?? attributes["val"] ?? "") else { break }
            let pt = CGFloat(v) / 2.0
            if inDocDefaults && inRPrDefault { defaultFontSizePt = pt }
            else if inStyleRPr { current.fontSizePt = pt }
        case "w:b", "b":
            guard inStyleRPr else { break }
            current.bold = (attributes["w:val"] ?? attributes["val"]) != "0"
        case "w:i", "i":
            guard inStyleRPr else { break }
            current.italic = (attributes["w:val"] ?? attributes["val"]) != "0"
        case "w:color", "color":
            guard inStyleRPr else { break }
            let val = attributes["w:val"] ?? attributes["val"] ?? ""
            if val != "auto" && !val.isEmpty { current.color = val }
        case "w:spacing", "spacing":
            guard inStylePPr else { break }
            if let v = Int(attributes["w:before"] ?? attributes["before"] ?? "") {
                currentSpacingBefore = CGFloat(v) / 20.0
            }
            if let v = Int(attributes["w:after"] ?? attributes["after"] ?? "") {
                currentSpacingAfter = CGFloat(v) / 20.0
            }
        default: break
        }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName: String?) {
        switch elementName {
        case "w:docDefaults", "docDefaults": inDocDefaults = false
        case "w:rPrDefault", "rPrDefault":   inRPrDefault = false
        case "w:pPr", "pPr":                 inStylePPr = false
        case "w:rPr", "rPr":                 inStyleRPr = false
        case "w:style", "style":
            if let id = currentStyleId {
                current.spacingBeforePt = currentSpacingBefore
                current.spacingAfterPt  = currentSpacingAfter
                styles[id] = current
            }
            currentStyleId = nil
        default: break
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
    private let styles: OOXMLStylesParser
    private var listCounters: [String: Int] = [:]

    init(numbering: OOXMLNumberingParser, styles: OOXMLStylesParser) {
        self.numbering = numbering
        self.styles = styles
    }

    // MARK: - Paragraph state
    private var inParagraph = false
    private var currentParaStyle = "Normal"
    private var currentParaSpacingAfter: CGFloat = 0
    private var currentParaSpacingBefore: CGFloat = 0
    private var currentParaAlignment: CTTextAlignment = .natural
    private var currentParaLeftIndent: CGFloat = 0
    private var currentParaFirstLineIndent: CGFloat = 0
    private var currentParaBackground: String? = nil
    private var currentParaBorderTopHex: String? = nil
    private var currentParaBorderTopWidth: CGFloat = 0
    private var currentParaBorderBottomHex: String? = nil
    private var currentParaBorderBottomWidth: CGFloat = 0
    private var currentRuns: [WordRunContent] = []
    private var inPBdr = false

    // MARK: - List state
    private var inNumPr = false
    private var currentListNumId: Int? = nil
    private var currentListIlvl: Int = 0

    // MARK: - Run state
    private var inRun = false
    private var currentRunBold = false
    private var currentRunItalic = false
    private var currentRunFontSize: CGFloat = 10
    private var currentRunColor: String? = nil
    private var currentRunUnderline = false
    private var currentRunStrikethrough = false
    private var currentRunFontFamily: String? = nil
    private var currentRunText = ""
    private var inText = false

    // MARK: - Table state
    private var inTable = false
    private var inRow = false
    private var inTrPr = false
    private var currentRowIsHeader = false
    private var inCell = false
    private var inTcPr = false
    private var inTcMar = false
    private var inTcBorders = false
    private var inTblGrid = false
    private var currentTableRows: [WordTableRow] = []
    private var currentRowCells: [WordTableCell] = []
    private var currentCellParagraphs: [WordParagraphContent] = []
    private var currentCellBackground: String? = nil
    private var currentCellBorderColorHex: String? = nil
    private var currentCellMarginTop: CGFloat? = nil
    private var currentCellMarginBottom: CGFloat? = nil
    private var currentCellMarginLeft: CGFloat? = nil
    private var currentCellMarginRight: CGFloat? = nil
    private var currentTableColWidths: [CGFloat] = []

    // MARK: - didStartElement

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName: String?,
                attributes: [String: String] = [:]) {
        switch elementName {

        // Table structure
        case "w:tbl", "tbl":
            inTable = true
            currentTableRows = []
            currentTableColWidths = []

        case "w:tblGrid", "tblGrid":
            inTblGrid = true

        case "w:gridCol", "gridCol":
            guard inTblGrid else { return }
            if let wStr = attributes["w:w"] ?? attributes["w"], let twips = Int(wStr) {
                currentTableColWidths.append(CGFloat(twips) / 20.0)
            }

        case "w:tr", "tr":
            guard inTable else { return }
            inRow = true
            inTrPr = false
            currentRowIsHeader = false
            currentRowCells = []

        case "w:trPr", "trPr":
            guard inRow else { return }
            inTrPr = true

        case "w:tblHeader", "tblHeader":
            guard inTrPr else { return }
            currentRowIsHeader = true

        case "w:tc", "tc":
            guard inTable, inRow else { return }
            inCell = true
            currentCellParagraphs = []
            currentCellBackground = nil
            currentCellBorderColorHex = nil
            currentCellMarginTop = nil
            currentCellMarginBottom = nil
            currentCellMarginLeft = nil
            currentCellMarginRight = nil

        case "w:tcPr", "tcPr":
            guard inCell else { return }
            inTcPr = true

        case "w:tcMar", "tcMar":
            guard inTcPr else { return }
            inTcMar = true

        case "w:tcBorders", "tcBorders":
            guard inTcPr else { return }
            inTcBorders = true

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
            currentParaBorderTopHex = nil
            currentParaBorderTopWidth = 0
            currentParaBorderBottomHex = nil
            currentParaBorderBottomWidth = 0
            currentRuns = []
            currentListNumId = nil
            currentListIlvl = 0

        case "w:pBdr", "pBdr":
            inPBdr = true

        case "w:r", "r":
            guard inParagraph else { return }
            inRun = true
            let sd = styles.resolve(styleId: currentParaStyle)
            currentRunBold = sd?.bold ?? false
            currentRunItalic = sd?.italic ?? false
            currentRunFontSize = sd?.fontSizePt ?? styles.defaultFontSizePt
            currentRunColor = sd?.color
            currentRunFontFamily = sd?.fontFamily ?? styles.defaultFontFamily
            currentRunUnderline = false
            currentRunStrikethrough = false
            currentRunText = ""

        // Paragraph properties
        case "w:pStyle", "pStyle":
            let val = attributes["w:val"] ?? attributes["val"] ?? "Normal"
            currentParaStyle = val
            if let sd = styles.resolve(styleId: val) {
                if let before = sd.spacingBeforePt { currentParaSpacingBefore = before }
                if let after  = sd.spacingAfterPt  { currentParaSpacingAfter  = after  }
            }

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
            if let hangStr = attributes["w:hanging"] ?? attributes["hanging"],
               let hangTwips = Int(hangStr) {
                currentParaFirstLineIndent = -CGFloat(hangTwips) / 20.0
            }

        case "w:shd", "shd":
            let val  = attributes["w:val"]   ?? attributes["val"]   ?? ""
            let fill = attributes["w:fill"]  ?? attributes["fill"]  ?? ""
            let shdColor = attributes["w:color"] ?? attributes["color"] ?? ""
            let bgHex = (val == "solid") ? shdColor : fill
            guard !bgHex.isEmpty && bgHex != "auto" && bgHex.uppercased() != "FFFFFF" else { break }
            if inTcPr {
                currentCellBackground = bgHex
            } else {
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

        // Context-sensitive: pBdr sides / tcMar sides / tcBorders sides
        case "w:top", "top":
            if inPBdr {
                let color = attributes["w:color"] ?? attributes["color"] ?? ""
                let sz = CGFloat(Int(attributes["w:sz"] ?? attributes["sz"] ?? "0") ?? 0) / 8.0
                currentParaBorderTopHex = (color.isEmpty || color == "auto") ? nil : color
                currentParaBorderTopWidth = sz
            } else if inTcMar {
                if let v = Int(attributes["w:w"] ?? attributes["w"] ?? "") {
                    currentCellMarginTop = CGFloat(v) / 20.0
                }
            } else if inTcBorders {
                let color = attributes["w:color"] ?? attributes["color"] ?? ""
                if !color.isEmpty && color != "auto" { currentCellBorderColorHex = color }
            }

        case "w:bottom", "bottom":
            if inPBdr {
                let color = attributes["w:color"] ?? attributes["color"] ?? ""
                let sz = CGFloat(Int(attributes["w:sz"] ?? attributes["sz"] ?? "0") ?? 0) / 8.0
                currentParaBorderBottomHex = (color.isEmpty || color == "auto") ? nil : color
                currentParaBorderBottomWidth = sz
            } else if inTcMar {
                if let v = Int(attributes["w:w"] ?? attributes["w"] ?? "") {
                    currentCellMarginBottom = CGFloat(v) / 20.0
                }
            } else if inTcBorders {
                let color = attributes["w:color"] ?? attributes["color"] ?? ""
                if !color.isEmpty && color != "auto" { currentCellBorderColorHex = color }
            }

        case "w:left", "left":
            if inTcMar {
                if let v = Int(attributes["w:w"] ?? attributes["w"] ?? "") {
                    currentCellMarginLeft = CGFloat(v) / 20.0
                }
            } else if inTcBorders {
                let color = attributes["w:color"] ?? attributes["color"] ?? ""
                if !color.isEmpty && color != "auto" { currentCellBorderColorHex = color }
            }

        case "w:right", "right":
            if inTcMar {
                if let v = Int(attributes["w:w"] ?? attributes["w"] ?? "") {
                    currentCellMarginRight = CGFloat(v) / 20.0
                }
            } else if inTcBorders {
                let color = attributes["w:color"] ?? attributes["color"] ?? ""
                if !color.isEmpty && color != "auto" { currentCellBorderColorHex = color }
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

        case "w:u", "u":
            guard inRun else { return }
            let val = attributes["w:val"] ?? attributes["val"] ?? "single"
            currentRunUnderline = (val != "none")

        case "w:strike", "strike":
            guard inRun else { return }
            let val = attributes["w:val"] ?? attributes["val"]
            currentRunStrikethrough = (val != "0")

        case "w:rFonts", "rFonts":
            guard inRun else { return }
            let family = attributes["w:ascii"] ?? attributes["ascii"]
                      ?? attributes["w:hAnsi"] ?? attributes["hAnsi"]
            if let f = family, !f.isEmpty { currentRunFontFamily = f }

        case "w:t", "t":
            if inRun { inText = true }

        case "w:br", "br":
            let breakType = attributes["w:type"] ?? attributes["type"]
            guard breakType == "page", inParagraph, !inCell else { return }
            if !currentRuns.isEmpty {
                elements.append(.paragraph(makeParagraph()))
                currentRuns = []
            }
            elements.append(.paragraph(WordParagraphContent(
                runs: [], styleName: "__pagebreak__", spacingAfterPt: 0)))

        case "w:sectPr", "sectPr":
            guard inParagraph, !inCell else { break }
            if !currentRuns.isEmpty {
                elements.append(.paragraph(makeParagraph()))
                currentRuns = []
            }
            elements.append(.paragraph(WordParagraphContent(
                runs: [], styleName: "__pagebreak__", spacingAfterPt: 0)))

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

    // MARK: - didEndElement

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName: String?) {
        switch elementName {

        case "w:tbl", "tbl":
            guard inTable else { return }
            elements.append(.table(WordTableContent(
                rows: currentTableRows,
                columnWidthsPt: currentTableColWidths)))
            currentTableRows = []
            inTable = false

        case "w:tblGrid", "tblGrid":
            inTblGrid = false

        case "w:trPr", "trPr":
            inTrPr = false

        case "w:tr", "tr":
            guard inTable, inRow else { return }
            // Use w:tblHeader flag if present; fall back to first-row detection
            let isHeader = currentRowIsHeader || currentTableRows.isEmpty
            currentTableRows.append(WordTableRow(cells: currentRowCells, isHeader: isHeader))
            currentRowCells = []
            inRow = false

        case "w:tcPr", "tcPr":
            inTcPr = false

        case "w:tcMar", "tcMar":
            inTcMar = false

        case "w:tcBorders", "tcBorders":
            inTcBorders = false

        case "w:tc", "tc":
            guard inTable, inCell else { return }
            let margins: WordTableCellMargins?
            if let t = currentCellMarginTop, let b = currentCellMarginBottom,
               let l = currentCellMarginLeft, let r = currentCellMarginRight {
                margins = WordTableCellMargins(top: t, bottom: b, left: l, right: r)
            } else {
                margins = nil
            }
            currentRowCells.append(WordTableCell(
                paragraphs: currentCellParagraphs,
                backgroundHex: currentCellBackground,
                borderColorHex: currentCellBorderColorHex,
                margins: margins))
            currentCellParagraphs = []
            currentCellBackground = nil
            inCell = false

        case "w:pBdr", "pBdr":
            inPBdr = false

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
                    hexColor: currentRunColor,
                    underline: currentRunUnderline,
                    strikethrough: currentRunStrikethrough,
                    fontFamily: currentRunFontFamily
                ))
            }
            inRun = false

        case "w:p", "p":
            guard inParagraph else { return }
            let para = makeParagraph()
            if inCell {
                currentCellParagraphs.append(para)
            } else {
                elements.append(.paragraph(para))
            }
            inParagraph = false

        default:
            break
        }
    }

    // MARK: - Private helpers

    private func makeParagraph() -> WordParagraphContent {
        let prefix = resolveListPrefix()
        var leftIndent = currentParaLeftIndent
        var firstLineIndent = currentParaFirstLineIndent
        // Apply numbering-level indent when paragraph has numPr but no inline w:ind
        if let numId = currentListNumId, leftIndent == 0 {
            let ilvl = currentListIlvl
            if let lvlLeft = numbering.leftIndent(numId: numId, ilvl: ilvl) {
                leftIndent = lvlLeft
            }
            if let lvlHang = numbering.hangingIndent(numId: numId, ilvl: ilvl),
               firstLineIndent == 0 {
                firstLineIndent = -lvlHang
            }
        }
        return WordParagraphContent(
            runs: currentRuns,
            styleName: currentParaStyle,
            spacingAfterPt: currentParaSpacingAfter,
            alignment: currentParaAlignment,
            listPrefix: prefix,
            spacingBeforePt: currentParaSpacingBefore,
            leftIndentPt: leftIndent,
            firstLineIndentPt: firstLineIndent,
            backgroundHex: currentParaBackground,
            borderTopHex: currentParaBorderTopHex,
            borderTopWidthPt: currentParaBorderTopWidth,
            borderBottomHex: currentParaBorderBottomHex,
            borderBottomWidthPt: currentParaBorderBottomWidth
        )
    }

    private func resolveListPrefix() -> String? {
        guard let numId = currentListNumId else { return nil }
        let ilvl = currentListIlvl
        guard let format = numbering.format(numId: numId, ilvl: ilvl) else { return nil }

        let key = "\(numId)-\(ilvl)"
        let counter = (listCounters[key] ?? (numbering.startVal(numId: numId, ilvl: ilvl) - 1)) + 1
        listCounters[key] = counter

        // Use a tab character after the glyph/number so the text jumps to the leftIndentPt
        // tab stop defined in the CTParagraphStyle. This aligns first-line text with the
        // wrapped lines (both start at leftIndentPt), matching Word's native rendering.
        switch format {
        case "bullet":
            let char = numbering.lvlText(numId: numId, ilvl: ilvl) ?? (ilvl == 0 ? "•" : "◦")
            return char + "\t"
        case "decimal":
            return "\(counter).\t"
        case "lowerLetter":
            return "\(letterLabel(counter - 1)).\t"
        case "upperLetter":
            return "\(letterLabel(counter - 1).uppercased()).\t"
        case "lowerRoman":
            return "\(romanNumeral(counter).lowercased()).\t"
        case "upperRoman":
            return "\(romanNumeral(counter)).\t"
        default:
            return "•\t"
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
