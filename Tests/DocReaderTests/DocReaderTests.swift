import XCTest
@testable import DocReader

final class DocReaderTests: XCTestCase {
    // MARK: - isSupported

    func testIsSupportedForAllSixExtensions() {
        let extensions = ["docx", "xlsx", "pptx", "doc", "xls", "ppt"]
        for ext in extensions {
            let url = URL(fileURLWithPath: "document.\(ext)")
            XCTAssertTrue(DocReader.isSupported(url: url), "\(ext) should be supported")
        }
    }

    func testIsNotSupportedForUnsupportedExtensions() {
        let unsupported = ["pdf", "txt", "rtf", "csv", "png", "jpg", ""]
        for ext in unsupported {
            let name = ext.isEmpty ? "noextension" : "file.\(ext)"
            let url = URL(fileURLWithPath: name)
            XCTAssertFalse(DocReader.isSupported(url: url), "\(ext) should not be supported")
        }
    }

    // MARK: - open (async throws)

    func testOpenThrowsFileNotFoundForMissingFile() async {
        let url = URL(fileURLWithPath: "/tmp/does_not_exist_12345.docx")
        await assertThrows(DocReaderError.fileNotFound) {
            _ = try await DocReader.open(url: url)
        }
    }

    func testOpenThrowsUnsupportedFormatForUnknownExtension() async throws {
        // Write a temp file with unsupported extension
        let tempURL = FileManager.default.temporaryDirectory
            .appendingPathComponent("test_unsupported.xyz")
        defer { try? FileManager.default.removeItem(at: tempURL) }
        try Data([0x00]).write(to: tempURL)

        await assertThrows(DocReaderError.unsupportedFormat) {
            _ = try await DocReader.open(url: tempURL)
        }
    }

    // MARK: - DocReaderError.errorDescription

    func testErrorDescriptionsAreNonNilEnglish() {
        let errors: [DocReaderError] = [
            .unsupportedFormat, .fileNotFound, .corruptedFile,
            .pageOutOfRange, .exportCancelled, .internalError("detail")
        ]
        for error in errors {
            XCTAssertNotNil(error.errorDescription,
                           "errorDescription should not be nil for \(error)")
        }
    }
}

// MARK: - Helpers

private func assertThrows<E: Error & Equatable>(
    _ expected: E,
    file: StaticString = #filePath,
    line: UInt = #line,
    _ block: () async throws -> Void
) async {
    do {
        try await block()
        XCTFail("Expected \(expected) to be thrown", file: file, line: line)
    } catch let error as E {
        XCTAssertEqual(error, expected, file: file, line: line)
    } catch {
        XCTFail("Unexpected error type: \(error)", file: file, line: line)
    }
}

extension DocReaderError: Equatable {
    public static func == (lhs: DocReaderError, rhs: DocReaderError) -> Bool {
        switch (lhs, rhs) {
        case (.unsupportedFormat, .unsupportedFormat): return true
        case (.fileNotFound, .fileNotFound): return true
        case (.corruptedFile, .corruptedFile): return true
        case (.pageOutOfRange, .pageOutOfRange): return true
        case (.exportCancelled, .exportCancelled): return true
        case (.internalError(let l), .internalError(let r)): return l == r
        default: return false
        }
    }
}
