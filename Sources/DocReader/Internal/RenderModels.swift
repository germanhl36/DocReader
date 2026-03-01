import Foundation
import CoreGraphics

// MARK: - Word content models

struct WordRunContent: Sendable {
    var text: String
    var bold: Bool
    var italic: Bool
    var fontSizePt: CGFloat
    var hexColor: String?
}

struct WordParagraphContent: Sendable {
    var runs: [WordRunContent]
    var styleName: String
    var spacingAfterPt: CGFloat
}

struct WordPageMargins: Sendable {
    var top: CGFloat
    var bottom: CGFloat
    var left: CGFloat
    var right: CGFloat
}

struct WordPageContent: Sendable {
    var paragraphs: [WordParagraphContent]
    var pageSize: CGSize
    var margins: WordPageMargins
}

// MARK: - Spreadsheet content models

struct SheetCellContent: Sendable {
    var col: Int
    var row: Int
    var text: String
}

struct SheetPageContent: Sendable {
    var sheetName: String
    var cells: [SheetCellContent]
}

// MARK: - Presentation content models

struct SlideTextBoxContent: Sendable {
    var frame: CGRect
    var lines: [String]
    var isTitle: Bool
}

struct SlidePageContent: Sendable {
    var textBoxes: [SlideTextBoxContent]
}

// MARK: - Page content enum

enum PageContent: Sendable {
    case word(WordPageContent)
    case sheet(SheetPageContent)
    case slide(SlidePageContent)
}
