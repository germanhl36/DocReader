import XCTest
@testable import DocReader

final class DocFormatDetectorTests: XCTestCase {
    // MARK: - Extension-based detection

    func testDetectsDocx() {
        let url = URL(fileURLWithPath: "file.docx")
        XCTAssertEqual(DocFormatDetector.detect(url: url), .docx)
    }

    func testDetectsXlsx() {
        let url = URL(fileURLWithPath: "file.xlsx")
        XCTAssertEqual(DocFormatDetector.detect(url: url), .xlsx)
    }

    func testDetectsPptx() {
        let url = URL(fileURLWithPath: "file.pptx")
        XCTAssertEqual(DocFormatDetector.detect(url: url), .pptx)
    }

    func testDetectsDoc() {
        let url = URL(fileURLWithPath: "file.doc")
        XCTAssertEqual(DocFormatDetector.detect(url: url), .doc)
    }

    func testDetectsXls() {
        let url = URL(fileURLWithPath: "file.xls")
        XCTAssertEqual(DocFormatDetector.detect(url: url), .xls)
    }

    func testDetectsPpt() {
        let url = URL(fileURLWithPath: "file.ppt")
        XCTAssertEqual(DocFormatDetector.detect(url: url), .ppt)
    }

    func testCaseInsensitiveExtension() {
        XCTAssertEqual(DocFormatDetector.format(forExtension: "DOCX"), .docx)
        XCTAssertEqual(DocFormatDetector.format(forExtension: "XlSx"), .xlsx)
    }

    func testUnknownExtensionReturnsNil() {
        XCTAssertNil(DocFormatDetector.format(forExtension: "pdf"))
        XCTAssertNil(DocFormatDetector.format(forExtension: "txt"))
        XCTAssertNil(DocFormatDetector.format(forExtension: ""))
        XCTAssertNil(DocFormatDetector.format(forExtension: "xyz"))
    }

    // MARK: - Magic bytes (OLE2)

    func testOLE2MagicBytesDetectedAsDoc() throws {
        let ole2Header: [UInt8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("test_ole2.doc")
        defer { try? FileManager.default.removeItem(at: tempURL) }

        var data = Data(ole2Header)
        data.append(contentsOf: Array(repeating: 0x00, count: 512))
        try data.write(to: tempURL)

        let detected = DocFormatDetector.detect(url: tempURL)
        XCTAssertEqual(detected?.family, .word)
    }

    func testZIPMagicBytesDetectedAsDocx() throws {
        let zipHeader: [UInt8] = [0x50, 0x4B, 0x03, 0x04]
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("test_zip.docx")
        defer { try? FileManager.default.removeItem(at: tempURL) }

        var data = Data(zipHeader)
        data.append(contentsOf: Array(repeating: 0x00, count: 512))
        try data.write(to: tempURL)

        let detected = DocFormatDetector.detect(url: tempURL)
        XCTAssertEqual(detected, .docx)
    }
}
