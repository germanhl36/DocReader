import XCTest
import DocReader

/// Integration tests for legacy (.doc, .xls, .ppt) parsers.
final class LegacyParserIntegrationTests: XCTestCase {
    private var fixturesURL: URL {
        Bundle.module.resourceURL!.appendingPathComponent("Fixtures")
    }

    // MARK: - DOC

    func testDocPageCount() async throws {
        let url = fixturesURL.appendingPathComponent("word_legacy_10page.doc")
        try XCTSkipUnless(FileManager.default.fileExists(atPath: url.path),
                          "Fixture word_legacy_10page.doc not found")

        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertGreaterThanOrEqual(count, 9)
        XCTAssertLessThanOrEqual(count, 11)
    }

    func testDocOpenPerformance() async throws {
        let url = fixturesURL.appendingPathComponent("word_legacy_10page.doc")
        try XCTSkipUnless(FileManager.default.fileExists(atPath: url.path),
                          "Fixture word_legacy_10page.doc not found")

        let start = Date()
        let doc = try await DocReader.open(url: url)
        _ = try await doc.pageCount
        let elapsed = Date().timeIntervalSince(start)
        XCTAssertLessThan(elapsed, 3.0, "Legacy .doc open should complete within 3 seconds")
    }

    // MARK: - XLS

    func testXlsSheetCount() async throws {
        let url = fixturesURL.appendingPathComponent("excel_legacy_3sheet.xls")
        try XCTSkipUnless(FileManager.default.fileExists(atPath: url.path),
                          "Fixture excel_legacy_3sheet.xls not found")

        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertGreaterThanOrEqual(count, 2)
        XCTAssertLessThanOrEqual(count, 4)
    }

    // MARK: - PPT

    func testPptSlideCount() async throws {
        let url = fixturesURL.appendingPathComponent("ppt_legacy_5slide.ppt")
        try XCTSkipUnless(FileManager.default.fileExists(atPath: url.path),
                          "Fixture ppt_legacy_5slide.ppt not found")

        let doc = try await DocReader.open(url: url)
        let count = try await doc.pageCount
        XCTAssertGreaterThanOrEqual(count, 4)
        XCTAssertLessThanOrEqual(count, 6)
    }
}
