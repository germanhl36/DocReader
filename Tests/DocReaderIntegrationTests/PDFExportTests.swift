import XCTest
import DocReader

/// Integration tests for PDF export across all six supported document formats.
///
/// Covers DOCR-31: verifies that ``DocReadable/exportPDF()`` and
/// ``DocReadable/exportPDF(pages:)`` produce valid PDF output for every
/// format family, and that out-of-range requests throw the correct error.
final class PDFExportTests: XCTestCase {

    private var fixturesURL: URL {
        Bundle.module.resourceURL!.appendingPathComponent("Fixtures")
    }

    /// Returns the fixture URL or skips the test if the file is absent.
    private func fixture(_ name: String) throws -> URL {
        let url = fixturesURL.appendingPathComponent(name)
        try XCTSkipUnless(
            FileManager.default.fileExists(atPath: url.path),
            "Fixture \(name) not found – run scripts/generate_fixtures.py first"
        )
        return url
    }

    /// Asserts that `data` begins with the `%PDF` magic bytes.
    private func assertValidPDF(_ data: Data, _ label: String) {
        XCTAssertGreaterThan(data.count, 0, "\(label): data must be non-empty")
        XCTAssertTrue(
            data.starts(with: [0x25, 0x50, 0x44, 0x46]),
            "\(label): output must start with %PDF"
        )
    }

    // MARK: - OOXML: DOCX

    func testDocxExportAllPages() async throws {
        let doc = try await DocReader.open(url: fixture("word_1page.docx"))
        let pdf = try await doc.exportPDF()
        assertValidPDF(pdf, "word_1page.docx – all pages")
    }

    func testDocxExportPageRange() async throws {
        let doc = try await DocReader.open(url: fixture("word_10page.docx"))
        let pdf = try await doc.exportPDF(pages: 0...2)
        assertValidPDF(pdf, "word_10page.docx – pages 0–2")
    }

    func testDocxExportLastPage() async throws {
        let doc = try await DocReader.open(url: fixture("word_10page.docx"))
        let count = try await doc.pageCount
        let pdf = try await doc.exportPDF(pages: (count - 1)...(count - 1))
        assertValidPDF(pdf, "word_10page.docx – last page")
    }

    // MARK: - OOXML: XLSX

    func testXlsxExportAllSheets() async throws {
        let doc = try await DocReader.open(url: fixture("excel_3sheet.xlsx"))
        let pdf = try await doc.exportPDF()
        assertValidPDF(pdf, "excel_3sheet.xlsx – all sheets")
    }

    func testXlsxExportSingleSheet() async throws {
        let doc = try await DocReader.open(url: fixture("excel_3sheet.xlsx"))
        let pdf = try await doc.exportPDF(pages: 1...1)
        assertValidPDF(pdf, "excel_3sheet.xlsx – sheet 2")
    }

    // MARK: - OOXML: PPTX

    func testPptxExportAllSlides() async throws {
        let doc = try await DocReader.open(url: fixture("ppt_5slide.pptx"))
        let pdf = try await doc.exportPDF()
        assertValidPDF(pdf, "ppt_5slide.pptx – all slides")
    }

    func testPptxExportSlideRange() async throws {
        let doc = try await DocReader.open(url: fixture("ppt_5slide.pptx"))
        let pdf = try await doc.exportPDF(pages: 1...3)
        assertValidPDF(pdf, "ppt_5slide.pptx – slides 2–4")
    }

    // MARK: - Legacy: DOC

    func testDocExportAllPages() async throws {
        let doc = try await DocReader.open(url: fixture("word_legacy_10page.doc"))
        let pdf = try await doc.exportPDF()
        assertValidPDF(pdf, "word_legacy_10page.doc – all pages")
    }

    func testDocExportPageRange() async throws {
        let doc = try await DocReader.open(url: fixture("word_legacy_10page.doc"))
        let pdf = try await doc.exportPDF(pages: 0...4)
        assertValidPDF(pdf, "word_legacy_10page.doc – pages 0–4")
    }

    // MARK: - Legacy: XLS

    func testXlsExportAllSheets() async throws {
        let doc = try await DocReader.open(url: fixture("excel_legacy_3sheet.xls"))
        let pdf = try await doc.exportPDF()
        assertValidPDF(pdf, "excel_legacy_3sheet.xls – all sheets")
    }

    func testXlsExportSingleSheet() async throws {
        let doc = try await DocReader.open(url: fixture("excel_legacy_3sheet.xls"))
        let pdf = try await doc.exportPDF(pages: 0...0)
        assertValidPDF(pdf, "excel_legacy_3sheet.xls – sheet 1")
    }

    // MARK: - Legacy: PPT

    func testPptExportAllSlides() async throws {
        let doc = try await DocReader.open(url: fixture("ppt_legacy_5slide.ppt"))
        let pdf = try await doc.exportPDF()
        assertValidPDF(pdf, "ppt_legacy_5slide.ppt – all slides")
    }

    func testPptExportSingleSlide() async throws {
        let doc = try await DocReader.open(url: fixture("ppt_legacy_5slide.ppt"))
        let pdf = try await doc.exportPDF(pages: 0...0)
        assertValidPDF(pdf, "ppt_legacy_5slide.ppt – slide 1")
    }

    // MARK: - Error handling

    func testExportPageOutOfRangeThrows() async throws {
        let doc = try await DocReader.open(url: fixture("word_1page.docx"))
        do {
            _ = try await doc.exportPDF(pages: 99...100)
            XCTFail("Expected pageOutOfRange to be thrown")
        } catch DocReaderError.pageOutOfRange {
            // expected
        }
    }

    func testExportedPDFGrowsWithPageCount() async throws {
        let doc = try await DocReader.open(url: fixture("word_10page.docx"))
        let count = try await doc.pageCount
        guard count >= 3 else { throw XCTSkip("Need at least 3 pages") }

        let pdfOne  = try await doc.exportPDF(pages: 0...0)
        let pdfFull = try await doc.exportPDF(pages: 0...(count - 1))
        XCTAssertLessThan(
            pdfOne.count, pdfFull.count,
            "A single-page export must be smaller than the full document export"
        )
    }
}
