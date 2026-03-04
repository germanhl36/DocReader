import XCTest
import CoreGraphics
import CoreText
@testable import DocReader

final class RichFormattingTests: XCTestCase {

    // MARK: - Alignment

    func testAlignmentCapture() {
        let para = WordParagraphContent(
            runs: [],
            styleName: "Normal",
            spacingAfterPt: 0,
            alignment: .center
        )
        XCTAssertEqual(para.alignment, .center)
    }

    func testAlignmentDefaultIsNatural() {
        let para = WordParagraphContent(runs: [], styleName: "Normal", spacingAfterPt: 0)
        XCTAssertEqual(para.alignment, .natural)
    }

    func testAlignmentJustified() {
        let para = WordParagraphContent(
            runs: [],
            styleName: "Normal",
            spacingAfterPt: 0,
            alignment: .justified
        )
        XCTAssertEqual(para.alignment, .justified)
    }

    // MARK: - List prefix

    func testListPrefixBullet() {
        let para = WordParagraphContent(
            runs: [],
            styleName: "Normal",
            spacingAfterPt: 0,
            listPrefix: "● "
        )
        XCTAssertEqual(para.listPrefix, "● ")
    }

    func testListPrefixCustomBulletChar() {
        let para = WordParagraphContent(
            runs: [],
            styleName: "ListParagraph",
            spacingAfterPt: 0,
            listPrefix: "■ "
        )
        XCTAssertEqual(para.listPrefix, "■ ")
    }

    func testListIndentFromNumPr() {
        let para = WordParagraphContent(
            runs: [],
            styleName: "ListParagraph",
            spacingAfterPt: 0,
            leftIndentPt: 36,
            firstLineIndentPt: -18
        )
        XCTAssertEqual(para.leftIndentPt, 36)
        XCTAssertEqual(para.firstLineIndentPt, -18)
    }

    func testListPrefixNumbered() {
        let para = WordParagraphContent(
            runs: [],
            styleName: "Normal",
            spacingAfterPt: 0,
            listPrefix: "1. "
        )
        XCTAssertEqual(para.listPrefix, "1. ")
    }

    func testListPrefixDefaultNil() {
        let para = WordParagraphContent(runs: [], styleName: "Normal", spacingAfterPt: 0)
        XCTAssertNil(para.listPrefix)
    }

    // MARK: - Table model construction

    func testTableModelConstruction() {
        let cell = WordTableCell(paragraphs: [])
        let row1 = WordTableRow(cells: [cell, cell, cell], isHeader: true)
        let row2 = WordTableRow(cells: [cell, cell, cell], isHeader: false)
        let table = WordTableContent(rows: [row1, row2])

        XCTAssertEqual(table.rows.count, 2)
        XCTAssertEqual(table.rows[0].cells.count, 3)
        XCTAssertTrue(table.rows[0].isHeader)
        XCTAssertFalse(table.rows[1].isHeader)
    }

    func testTableCellContainsParagraphs() {
        let run = WordRunContent(text: "cell text", bold: false, italic: false, fontSizePt: 12, hexColor: nil)
        let para = WordParagraphContent(runs: [run], styleName: "Normal", spacingAfterPt: 0)
        let cell = WordTableCell(paragraphs: [para])
        XCTAssertEqual(cell.paragraphs.count, 1)
        XCTAssertEqual(cell.paragraphs[0].runs[0].text, "cell text")
    }

    func testTableCellBackgroundDefaultNil() {
        let cell = WordTableCell(paragraphs: [])
        XCTAssertNil(cell.backgroundHex)
    }

    func testTableCellBackground() {
        let cell = WordTableCell(paragraphs: [], backgroundHex: "1E3A5F")
        XCTAssertEqual(cell.backgroundHex, "1E3A5F")
    }

    // MARK: - WordElement enum

    func testWordElementParagraphCase() {
        let para = WordParagraphContent(runs: [], styleName: "Normal", spacingAfterPt: 0)
        let element = WordElement.paragraph(para)
        if case .paragraph(let p) = element {
            XCTAssertEqual(p.styleName, "Normal")
        } else {
            XCTFail("Expected .paragraph case")
        }
    }

    func testWordElementTableCase() {
        let table = WordTableContent(rows: [])
        let element = WordElement.table(table)
        if case .table(let t) = element {
            XCTAssertEqual(t.rows.count, 0)
        } else {
            XCTFail("Expected .table case")
        }
    }

    // MARK: - splitIntoWordPages with mixed elements

    func testElementSplitOnPageBreak() {
        let run = WordRunContent(text: "text", bold: false, italic: false, fontSizePt: 12, hexColor: nil)
        let para = WordParagraphContent(runs: [run], styleName: "Normal", spacingAfterPt: 0)
        let pageBreak = WordParagraphContent(runs: [], styleName: "__pagebreak__", spacingAfterPt: 0)
        let table = WordTableContent(rows: [WordTableRow(cells: [], isHeader: false)])

        let elements: [WordElement] = [
            .paragraph(para),
            .table(table),
            .paragraph(pageBreak),
            .paragraph(para),
        ]

        let pages = splitIntoWordPages(
            elements: elements,
            pageSize: CGSize(width: 612, height: 792),
            margins: WordPageMargins(top: 72, bottom: 72, left: 72, right: 72)
        )

        XCTAssertEqual(pages.count, 2)
        XCTAssertEqual(pages[0].elements.count, 2, "First page: paragraph + table")
        XCTAssertEqual(pages[1].elements.count, 1, "Second page: paragraph")
    }

    func testTablePreservedAcrossPageSplit() {
        let table = WordTableContent(rows: [
            WordTableRow(cells: [WordTableCell(paragraphs: [])], isHeader: true)
        ])
        let pageBreak = WordParagraphContent(runs: [], styleName: "__pagebreak__", spacingAfterPt: 0)

        let elements: [WordElement] = [
            .table(table),
            .paragraph(pageBreak),
            .table(table),
        ]

        let pages = splitIntoWordPages(
            elements: elements,
            pageSize: CGSize(width: 612, height: 792),
            margins: WordPageMargins(top: 72, bottom: 72, left: 72, right: 72)
        )

        XCTAssertEqual(pages.count, 2)
        if case .table(let t) = pages[0].elements[0] {
            XCTAssertEqual(t.rows.count, 1)
        } else {
            XCTFail("Expected .table on first page")
        }
    }

    // MARK: - Spacing and indent fields

    func testSpacingBeforeDefault() {
        let para = WordParagraphContent(runs: [], styleName: "Normal", spacingAfterPt: 0)
        XCTAssertEqual(para.spacingBeforePt, 0)
    }

    func testIndentFields() {
        let para = WordParagraphContent(
            runs: [],
            styleName: "Normal",
            spacingAfterPt: 0,
            leftIndentPt: 36,
            firstLineIndentPt: -18
        )
        XCTAssertEqual(para.leftIndentPt, 36)
        XCTAssertEqual(para.firstLineIndentPt, -18)
    }

    func testBackgroundHex() {
        let para = WordParagraphContent(
            runs: [],
            styleName: "Normal",
            spacingAfterPt: 0,
            backgroundHex: "F0F0F0"
        )
        XCTAssertEqual(para.backgroundHex, "F0F0F0")
    }

    // MARK: - Run decorations and font family

    func testRunUnderlineDefaultFalse() {
        let run = WordRunContent(text: "t", bold: false, italic: false,
                                 fontSizePt: 12, hexColor: nil)
        XCTAssertFalse(run.underline)
    }

    func testRunUnderlineExplicit() {
        let run = WordRunContent(text: "t", bold: false, italic: false,
                                 fontSizePt: 12, hexColor: nil, underline: true)
        XCTAssertTrue(run.underline)
    }

    func testRunStrikethroughDefaultFalse() {
        let run = WordRunContent(text: "t", bold: false, italic: false,
                                 fontSizePt: 12, hexColor: nil)
        XCTAssertFalse(run.strikethrough)
    }

    func testRunStrikethroughExplicit() {
        let run = WordRunContent(text: "t", bold: false, italic: false,
                                 fontSizePt: 12, hexColor: nil, strikethrough: true)
        XCTAssertTrue(run.strikethrough)
    }

    func testRunFontFamilyDefaultNil() {
        let run = WordRunContent(text: "t", bold: false, italic: false,
                                 fontSizePt: 12, hexColor: nil)
        XCTAssertNil(run.fontFamily)
    }

    func testRunFontFamilyExplicit() {
        let run = WordRunContent(text: "t", bold: false, italic: false,
                                 fontSizePt: 12, hexColor: nil, fontFamily: "Calibri")
        XCTAssertEqual(run.fontFamily, "Calibri")
    }

    // MARK: - Paragraph border fields

    func testParagraphBorderTopDefaultNil() {
        let para = WordParagraphContent(runs: [], styleName: "Normal", spacingAfterPt: 0)
        XCTAssertNil(para.borderTopHex)
        XCTAssertEqual(para.borderTopWidthPt, 0)
    }

    func testParagraphBorderBottomExplicit() {
        let para = WordParagraphContent(
            runs: [], styleName: "Heading1", spacingAfterPt: 0,
            borderBottomHex: "2E86C1", borderBottomWidthPt: 1)
        XCTAssertEqual(para.borderBottomHex, "2E86C1")
        XCTAssertEqual(para.borderBottomWidthPt, 1)
    }

    // MARK: - Table cell margins

    func testTableCellMarginsDefaultNil() {
        let cell = WordTableCell(paragraphs: [])
        XCTAssertNil(cell.margins)
    }

    func testTableCellMarginsExplicit() {
        let margins = WordTableCellMargins(top: 6, bottom: 6, left: 9, right: 9)
        let cell = WordTableCell(paragraphs: [], margins: margins)
        XCTAssertEqual(cell.margins?.left, 9)
        XCTAssertEqual(cell.margins?.top, 6)
    }

    func testTableCellBorderColorDefaultNil() {
        let cell = WordTableCell(paragraphs: [])
        XCTAssertNil(cell.borderColorHex)
    }

    func testTableCellBorderColorExplicit() {
        let cell = WordTableCell(paragraphs: [], borderColorHex: "AAAAAA")
        XCTAssertEqual(cell.borderColorHex, "AAAAAA")
    }

    // MARK: - Table column widths

    func testTableColumnWidthsDefaultEmpty() {
        let table = WordTableContent(rows: [])
        XCTAssertTrue(table.columnWidthsPt.isEmpty)
    }

    func testTableColumnWidthsExplicit() {
        let table = WordTableContent(rows: [], columnWidthsPt: [234, 234])
        XCTAssertEqual(table.columnWidthsPt, [234, 234])
    }
}
