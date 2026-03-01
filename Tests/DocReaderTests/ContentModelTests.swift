import XCTest
import CoreGraphics
@testable import DocReader

final class ContentModelTests: XCTestCase {

    // MARK: - Struct field access

    func testWordRunContentDefaults() {
        let run = WordRunContent(text: "hello", bold: false, italic: false, fontSizePt: 12, hexColor: nil)
        XCTAssertEqual(run.text, "hello")
        XCTAssertFalse(run.bold)
        XCTAssertFalse(run.italic)
        XCTAssertEqual(run.fontSizePt, 12)
        XCTAssertNil(run.hexColor)
    }

    func testWordParagraphContentFields() {
        let run = WordRunContent(text: "text", bold: true, italic: false, fontSizePt: 14, hexColor: "000000")
        let para = WordParagraphContent(runs: [run], styleName: "Heading1", spacingAfterPt: 6)
        XCTAssertEqual(para.runs.count, 1)
        XCTAssertEqual(para.styleName, "Heading1")
        XCTAssertEqual(para.spacingAfterPt, 6)
    }

    // MARK: - Column label conversion

    func testColumnLabelConversion() {
        XCTAssertEqual(columnLabel(0), "A")
        XCTAssertEqual(columnLabel(25), "Z")
        XCTAssertEqual(columnLabel(26), "AA")
        XCTAssertEqual(columnLabel(51), "AZ")
        XCTAssertEqual(columnLabel(52), "BA")
    }

    // MARK: - Column index from cell reference

    func testColumnIndexFromRef() {
        XCTAssertEqual(columnIndex(from: "A1"), 0)
        XCTAssertEqual(columnIndex(from: "B3"), 1)
        XCTAssertEqual(columnIndex(from: "Z1"), 25)
        XCTAssertEqual(columnIndex(from: "AA1"), 26)
    }

    func testColumnLabelRoundTrip() {
        for i in 0..<100 {
            let label = columnLabel(i)
            let index = columnIndex(from: label + "1")
            XCTAssertEqual(index, i, "Round-trip failed for index \(i) → \(label)")
        }
    }

    // MARK: - Page break splitting

    func testPageBreakSplitting() {
        let run = WordRunContent(text: "text", bold: false, italic: false, fontSizePt: 12, hexColor: nil)
        let paragraphs: [WordParagraphContent] = [
            WordParagraphContent(runs: [run], styleName: "Normal", spacingAfterPt: 0),
            WordParagraphContent(runs: [], styleName: "__pagebreak__", spacingAfterPt: 0),
            WordParagraphContent(runs: [run], styleName: "Normal", spacingAfterPt: 0),
        ]
        let pages = splitIntoWordPages(
            paragraphs: paragraphs,
            pageSize: CGSize(width: 612, height: 792),
            margins: WordPageMargins(top: 72, bottom: 72, left: 72, right: 72)
        )
        XCTAssertEqual(pages.count, 2)
        XCTAssertEqual(pages[0].paragraphs.count, 1)
        XCTAssertEqual(pages[1].paragraphs.count, 1)
    }

    func testPageBreakSplittingNoBreaks() {
        let run = WordRunContent(text: "text", bold: false, italic: false, fontSizePt: 12, hexColor: nil)
        let paragraphs: [WordParagraphContent] = [
            WordParagraphContent(runs: [run], styleName: "Normal", spacingAfterPt: 0),
            WordParagraphContent(runs: [run], styleName: "Normal", spacingAfterPt: 0),
        ]
        let pages = splitIntoWordPages(
            paragraphs: paragraphs,
            pageSize: CGSize(width: 612, height: 792),
            margins: WordPageMargins(top: 72, bottom: 72, left: 72, right: 72)
        )
        XCTAssertEqual(pages.count, 1)
        XCTAssertEqual(pages[0].paragraphs.count, 2)
    }

    func testPageBreakSplittingMultipleBreaks() {
        let run = WordRunContent(text: "p", bold: false, italic: false, fontSizePt: 12, hexColor: nil)
        let br  = WordParagraphContent(runs: [], styleName: "__pagebreak__", spacingAfterPt: 0)
        let pg  = WordParagraphContent(runs: [run], styleName: "Normal", spacingAfterPt: 0)
        let pages = splitIntoWordPages(
            paragraphs: [pg, br, pg, br, pg],
            pageSize: CGSize(width: 612, height: 792),
            margins: WordPageMargins(top: 72, bottom: 72, left: 72, right: 72)
        )
        XCTAssertEqual(pages.count, 3)
    }

    // MARK: - Hex color resolution

    func testHexColorResolution() {
        let color = WordPageRenderer.resolveColor("FF0000")
        XCTAssertNotNil(color)
        let components = color?.components ?? []
        XCTAssertGreaterThan(components.first ?? 0, 0.9, "Red component should be near 1")
    }

    func testHexColorBlack() {
        let color = WordPageRenderer.resolveColor("000000")
        XCTAssertNotNil(color)
    }

    func testHexColorWhite() {
        let color = WordPageRenderer.resolveColor("FFFFFF")
        XCTAssertNotNil(color)
        let components = color?.components ?? []
        XCTAssertEqual(components.count, 4)
        XCTAssertGreaterThan(components[0], 0.9)
        XCTAssertGreaterThan(components[1], 0.9)
        XCTAssertGreaterThan(components[2], 0.9)
    }

    func testHexColorInvalidReturnsNil() {
        XCTAssertNil(WordPageRenderer.resolveColor("ZZZ"))
        XCTAssertNil(WordPageRenderer.resolveColor("12345"))   // too short
        XCTAssertNil(WordPageRenderer.resolveColor("1234567")) // too long
    }

    func testHexColorWithHashPrefix() {
        let color = WordPageRenderer.resolveColor("#00FF00")
        XCTAssertNotNil(color)
    }
}
