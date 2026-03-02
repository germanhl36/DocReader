import XCTest
import DocReader

/// Integration tests for rich formatting features (Iteration 7).
///
/// Verifies that tables and list documents export valid PDFs,
/// and that the SDD regression still passes.
final class RichFormattingIntegrationTests: XCTestCase {

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

    // MARK: - Table document

    func testDocxWithTableExportsPDF() async throws {
        let url = try fixture("word_table.docx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertGreaterThan(pdf.count, 2_000, "Table DOCX PDF should be > 2 KB")
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]), "Must start with %PDF")
    }

    func testDocxWithTablePageCount() async throws {
        let url = try fixture("word_table.docx")
        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertEqual(count, 1, "Single-page table document")
    }

    // MARK: - Lists document

    func testDocxWithListsExportsPDF() async throws {
        let url = try fixture("word_lists.docx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertGreaterThan(pdf.count, 1_000, "Lists DOCX PDF should be > 1 KB")
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]), "Must start with %PDF")
    }

    func testDocxWithListsPageCount() async throws {
        let url = try fixture("word_lists.docx")
        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertEqual(count, 1, "Single-page lists document")
    }

    // MARK: - SDD regression

    func testDocxSddPageCountStillNine() async throws {
        let url = try fixture("DocReader_SDD.docx")
        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertEqual(count, 9, "DocReader_SDD.docx should still have 9 pages after Iteration 7")
    }

    func testDocxSddExportStillValid() async throws {
        let url = try fixture("DocReader_SDD.docx")
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        XCTAssertGreaterThan(pdf.count, 20_000, "SDD PDF should be > 20 KB")
        XCTAssertTrue(pdf.starts(with: [0x25, 0x50, 0x44, 0x46]), "Must start with %PDF")
    }
}
