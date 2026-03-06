import XCTest
@testable import DocReader

/// Unit tests for PrintExporter. These tests use a minimal 1-page PDF generated
/// programmatically (CoreGraphics) so no fixture files are required.
final class PrintExporterTests: XCTestCase {

    // MARK: - Minimal 1-page PDF fixture

    /// Generates a tiny single-page PDF (Letter, white background) in memory.
    private func makeMinimalPDF() throws -> Data {
        let pageSize = CGSize(width: 612, height: 792)
        var mediaBox = CGRect(origin: .zero, size: pageSize)
        let pdfData = NSMutableData()
        guard let consumer = CGDataConsumer(data: pdfData),
              let ctx = CGContext(consumer: consumer, mediaBox: &mediaBox, nil) else {
            throw DocReaderError.internalError("Could not create PDF context")
        }
        ctx.beginPDFPage(nil)
        ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
        ctx.fill(mediaBox)
        ctx.endPDFPage()
        ctx.closePDF()
        return pdfData as Data
    }

    // MARK: - PWG-Raster tests

    func testPWGRasterMagicBytes() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportPWGRaster(pdf: pdf, resolution: 72)
        XCTAssertEqual(Array(out.prefix(4)), Array("RaS2".utf8))
    }

    func testPWGRasterHeaderSize() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportPWGRaster(pdf: pdf, resolution: 72)
        // magic (4) + header (1796) = 1800 bytes minimum
        XCTAssertGreaterThanOrEqual(out.count, 1800)
    }

    func testPWGRasterResolutionField() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportPWGRaster(pdf: pdf, resolution: 300)
        let bytes = [UInt8](out)
        // HWResolutionX at offset 4 + 276 = 280
        let offset = 4 + 276
        let value = (UInt32(bytes[offset]) << 24)
                  | (UInt32(bytes[offset + 1]) << 16)
                  | (UInt32(bytes[offset + 2]) << 8)
                  |  UInt32(bytes[offset + 3])
        XCTAssertEqual(value, 300)
    }

    // MARK: - URF tests

    func testURFMagicBytes() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportURF(pdf: pdf, resolution: 72)
        XCTAssertEqual(Array(out.prefix(8)), Array("UNIRAST\0".utf8))
    }

    func testURFPageCount() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportURF(pdf: pdf, resolution: 72)
        let bytes = [UInt8](out)
        // Page count is uint32 BE at bytes [8..11]
        let count = (UInt32(bytes[8]) << 24)
                  | (UInt32(bytes[9]) << 16)
                  | (UInt32(bytes[10]) << 8)
                  |  UInt32(bytes[11])
        XCTAssertEqual(count, 1)
    }

    // MARK: - PackBits tests

    func testPackBitsRunEncoding() {
        // Identical scanline: all red pixels → compact run encoding
        let red: [UInt8] = Array(repeating: 0, count: 0)  // start with empty
        var row = [UInt8]()
        for _ in 0..<10 { row += [255, 0, 0] }            // 10 identical red pixels
        let encoded = PrintExporter.packBitsRow(row[...])
        // A 10-pixel identical run encodes as 2 bytes header + 3 pixel = 5 bytes total (for 10 pixels)
        // Much smaller than raw 30 bytes
        XCTAssertLessThan(encoded.count, row.count)
        _ = red  // suppress unused warning
    }

    func testPackBitsLiteralEncoding() {
        // All-different pixels → literal (non-run) path
        var row = [UInt8]()
        for i in 0..<10 { row += [UInt8(i * 25), UInt8(i * 10), UInt8(i)] }
        let encoded = PrintExporter.packBitsRow(row[...])
        // Must decode to something (non-empty)
        XCTAssertFalse(encoded.isEmpty)
        // First byte < 0x80 means literal run header
        XCTAssertLessThan(encoded[0], 0x80)
    }

    // MARK: - PCL 5 tests

    func testPCLMagicReset() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportPCL(pdf: pdf, resolution: 72)
        XCTAssertEqual(Array(out.prefix(2)), [0x1B, 0x45])  // ESC E
    }

    func testPCLContainsRasterStart() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportPCL(pdf: pdf, resolution: 72)
        let bytes = [UInt8](out)
        // Look for ESC*r0A  =  [0x1B, 0x2A, 0x72, 0x30, 0x41]
        let needle: [UInt8] = [0x1B, 0x2A, 0x72, 0x30, 0x41]
        let found = bytes.windows(ofCount: needle.count).contains { Array($0) == needle }
        XCTAssertTrue(found, "PCL output should contain ESC*r0A (start raster)")
    }

    // MARK: - PCL XL tests

    func testPCLXLStreamHeader() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportPCLXL(pdf: pdf, resolution: 72)
        let bytes = [UInt8](out)
        let needle = Array(") HP-PCL XL;3;0".utf8)
        let found = bytes.windows(ofCount: needle.count).contains { Array($0) == needle }
        XCTAssertTrue(found, "') HP-PCL XL;3;0' not found in PCL XL stream header")
    }

    func testPCLXLBeginSessionOp() async throws {
        let pdf = try makeMinimalPDF()
        let out = try await PrintExporter.exportPCLXL(pdf: pdf, resolution: 72)
        let bytes = [UInt8](out)
        // 0x41 = BeginSession opcode should appear after ASCII header
        XCTAssertTrue(bytes.contains(0x41), "BeginSession (0x41) opcode not found")
    }
}

// MARK: - Sliding window helper (avoids importing Algorithms)
private extension Array {
    func windows(ofCount n: Int) -> [[Element]] {
        guard n <= count else { return [] }
        return (0...(count - n)).map { Array(self[$0..<($0 + n)]) }
    }
}
