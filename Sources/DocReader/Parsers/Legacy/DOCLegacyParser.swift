import Foundation
import CoreGraphics
import OLEKit

/// Parses `.doc` files using the OLE2 Compound File Binary format via OLEKit.
actor DOCLegacyParser: DocReadable {
    nonisolated let url: URL
    nonisolated let format: DocumentFormat = .doc

    private var _pageCount: Int?
    private var _metadata: DocumentMetadata?

    init(url: URL) {
        self.url = url
    }

    var pageCount: Int {
        get throws {
            if let cached = _pageCount { return cached }
            let count = try parsePageCount()
            _pageCount = count
            return count
        }
    }

    func pageSize(at index: Int) throws -> CGSize {
        // Try to read from SummaryInformation stream; fall back to US Letter
        if let size = try? parsePageSize(), size.width > 0 {
            return size
        }
        return CGSize(width: 612, height: 792)
    }

    var metadata: DocumentMetadata {
        get throws {
            if let cached = _metadata { return cached }
            let meta = try parseMetadata()
            _metadata = meta
            return meta
        }
    }

    func exportPDF() async throws -> Data {
        let count = try pageCount
        guard count > 0 else { throw DocReaderError.corruptedFile }
        return try await exportPDF(pages: 0...(count - 1))
    }

    func exportPDF(pages: ClosedRange<Int>) async throws -> Data {
        let count = try pageCount
        guard pages.lowerBound >= 0, pages.upperBound < count else {
            throw DocReaderError.pageOutOfRange
        }
        return try await PDFExporter.export(parser: self, pages: pages)
    }

    // MARK: - Parsing

    private func parsePageCount() throws -> Int {
        let cfb = try openCFB()
        // Read \x05SummaryInformation stream for PIDSI_PAGECOUNT (0x0E)
        if let count = readSummaryPageCount(from: cfb) {
            return count
        }
        // Fallback: return 1
        return 1
    }

    private func parsePageSize() throws -> CGSize {
        let cfb = try openCFB()
        guard let sizeBytes = readSummaryPageSize(from: cfb) else {
            return CGSize(width: 612, height: 792)
        }
        // Values are in twips (1/20 of a point)
        let width = CGFloat(sizeBytes.0) / 20.0
        let height = CGFloat(sizeBytes.1) / 20.0
        guard width > 0, height > 0 else {
            return CGSize(width: 612, height: 792)
        }
        return CGSize(width: width, height: height)
    }

    private func parseMetadata() throws -> DocumentMetadata {
        let cfb = try openCFB()
        return readSummaryMetadata(from: cfb)
    }

    // MARK: - OLE helpers

    private func openCFB() throws -> OLEFile {
        do {
            return try OLEFile(url.path)
        } catch {
            throw DocReaderError.corruptedFile
        }
    }

    /// Reads PIDSI_PAGECOUNT from \x05SummaryInformation using OLEKit.
    private func readSummaryPageCount(from ole: OLEFile) -> Int? {
        guard let entry = ole.root.children.first(where: { $0.name == "\u{05}SummaryInformation" }),
              let reader = try? ole.stream(entry) else { return nil }
        let data = reader.readData(ofLength: reader.totalBytes)
        return OLEPropertySetReader.readInt(from: data, propertyID: 0x0E)
    }

    private func readSummaryPageSize(from ole: OLEFile) -> (Int, Int)? {
        guard let entry = ole.root.children.first(where: { $0.name == "\u{05}DocumentSummaryInformation" }),
              let reader = try? ole.stream(entry) else { return nil }
        let data = reader.readData(ofLength: reader.totalBytes)
        if let w = OLEPropertySetReader.readInt(from: data, propertyID: 0x13),
           let h = OLEPropertySetReader.readInt(from: data, propertyID: 0x14) {
            return (w, h)
        }
        return nil
    }

    private func readSummaryMetadata(from ole: OLEFile) -> DocumentMetadata {
        guard let entry = ole.root.children.first(where: { $0.name == "\u{05}SummaryInformation" }),
              let reader = try? ole.stream(entry) else { return DocumentMetadata() }
        let data = reader.readData(ofLength: reader.totalBytes)
        let title    = OLEPropertySetReader.readString(from: data, propertyID: 0x02)
        let author   = OLEPropertySetReader.readString(from: data, propertyID: 0x04)
        let created  = OLEPropertySetReader.readDate(from: data, propertyID: 0x0C)
        let modified = OLEPropertySetReader.readDate(from: data, propertyID: 0x0D)
        return DocumentMetadata(title: title, author: author,
                                modifiedDate: modified, createdDate: created)
    }
}
