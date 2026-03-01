import XCTest
import DocReader

/// Integration tests for real content rendering (Iteration 6).
///
/// Covers DOCR-39…DOCR-54: verifies that OOXML parsers extract real content
/// and that exported PDFs contain meaningful data beyond placeholder bytes.
final class ContentRenderingTests: XCTestCase {

    private var fixturesURL: URL {
        Bundle.module.resourceURL!.appendingPathComponent("Fixtures")
    }

    /// Returns the fixture URL or skips the test if the file is absent.
    private func fixture(_ name: String) throws -> URL {
        let url = fixturesURL.appendingPathComponent(name)
        try XCTSkipUnless(
            FileManager.default.fileExists(atPath: url.path),
            "Fixture \(name) not found – run scripts/generate_fixtures.py or add the file first"
        )
        return url
    }

    // MARK: - DOCX: SDD document (real-world fixture)

    /// Requires DocReader_SDD.docx to be placed in the Fixtures directory.
    func testDocxSddPageCount() async throws {
        let url = try fixture("DocReader_SDD.docx")
        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertEqual(count, 9, "DocReader_SDD.docx should have 9 pages (8 page breaks + 1)")
    }

    func testDocxSddExportProducesNonTrivialPDF() async throws {
        let url = try fixture("DocReader_SDD.docx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertGreaterThan(pdf.count, 20_000, "SDD PDF should be > 20 KB")
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]), "Must start with %PDF")
    }

    // MARK: - DOCX: multi-page fixture

    func testDocxMultiPageRenderSize() async throws {
        let url = try fixture("word_10page.docx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertGreaterThan(pdf.count, 5_000, "10-page DOCX PDF should be > 5 KB")
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]), "Must start with %PDF")
    }

    func testDocxSinglePageRenderSize() async throws {
        let url = try fixture("word_1page.docx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertGreaterThan(pdf.count, 500, "1-page DOCX PDF should be > 500 bytes")
    }

    // MARK: - XLSX: cell extraction

    func testXlsxCellsExtracted() async throws {
        let url = try fixture("excel_3sheet.xlsx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        // PDF with cell content should be larger than a blank placeholder
        XCTAssertGreaterThan(pdf.count, 1_000, "XLSX PDF with cells should be > 1 KB")
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]), "Must start with %PDF")
    }

    func testXlsxPageCountMatchesSheets() async throws {
        let url = try fixture("excel_3sheet.xlsx")
        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertEqual(count, 3, "excel_3sheet.xlsx should have 3 sheets")
    }

    // MARK: - PPTX: slide text extraction

    func testPptxSlideTextExtracted() async throws {
        let url = try fixture("ppt_5slide.pptx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertGreaterThan(pdf.count, 1_000, "PPTX PDF with text should be > 1 KB")
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]), "Must start with %PDF")
    }

    func testPptxPageCountMatchesSlides() async throws {
        let url = try fixture("ppt_5slide.pptx")
        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertEqual(count, 5, "ppt_5slide.pptx should have 5 slides")
    }

    // MARK: - Regression: pre-existing tests still export valid PDFs

    func testDocxExportStillValid() async throws {
        let url = try fixture("word_1page.docx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]))
    }

    func testXlsxExportStillValid() async throws {
        let url = try fixture("excel_3sheet.xlsx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]))
    }

    func testPptxExportStillValid() async throws {
        let url = try fixture("ppt_5slide.pptx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]))
    }
}
