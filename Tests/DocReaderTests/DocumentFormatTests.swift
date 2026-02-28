import XCTest
@testable import DocReader

final class DocumentFormatTests: XCTestCase {
    // MARK: - family

    func testWordFamily() {
        XCTAssertEqual(DocumentFormat.docx.family, .word)
        XCTAssertEqual(DocumentFormat.doc.family, .word)
    }

    func testSpreadsheetFamily() {
        XCTAssertEqual(DocumentFormat.xlsx.family, .spreadsheet)
        XCTAssertEqual(DocumentFormat.xls.family, .spreadsheet)
    }

    func testPresentationFamily() {
        XCTAssertEqual(DocumentFormat.pptx.family, .presentation)
        XCTAssertEqual(DocumentFormat.ppt.family, .presentation)
    }

    // MARK: - isLegacy / isOOXML

    func testLegacyFormats() {
        XCTAssertTrue(DocumentFormat.doc.isLegacy)
        XCTAssertTrue(DocumentFormat.xls.isLegacy)
        XCTAssertTrue(DocumentFormat.ppt.isLegacy)
    }

    func testOOXMLFormats() {
        XCTAssertTrue(DocumentFormat.docx.isOOXML)
        XCTAssertTrue(DocumentFormat.xlsx.isOOXML)
        XCTAssertTrue(DocumentFormat.pptx.isOOXML)
    }

    func testIsLegacyAndIsOOXMLAreMutuallyExclusive() {
        for format in DocumentFormat.allCases {
            XCTAssertNotEqual(format.isLegacy, format.isOOXML,
                              "\(format) must be either legacy or OOXML, not both")
        }
    }

    // MARK: - fileExtension

    func testFileExtensionMatchesRawValue() {
        for format in DocumentFormat.allCases {
            XCTAssertEqual(format.fileExtension, format.rawValue)
        }
    }

    // MARK: - CaseIterable

    func testAllCasesCount() {
        XCTAssertEqual(DocumentFormat.allCases.count, 6)
    }
}
