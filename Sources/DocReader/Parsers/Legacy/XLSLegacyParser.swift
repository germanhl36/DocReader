import Foundation
import CoreGraphics
import OLEKit

/// Parses `.xls` files (BIFF8 / OLE2 Compound File Binary).
actor XLSLegacyParser: DocReadable {
    nonisolated let url: URL
    nonisolated let format: DocumentFormat = .xls

    private var _pageCount: Int?
    private var _metadata: DocumentMetadata?

    init(url: URL) {
        self.url = url
    }

    var pageCount: Int {
        get throws {
            if let cached = _pageCount { return cached }
            let count = try parseSheetCount()
            _pageCount = count
            return count
        }
    }

    func pageSize(at index: Int) throws -> CGSize {
        // XLS defaults to US Letter (BIFF8 SETUP record parsing is complex)
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

    private func parseSheetCount() throws -> Int {
        let cfb: OLEFile
        do {
            cfb = try OLEFile(url.path)
        } catch {
            throw DocReaderError.corruptedFile
        }

        // The Workbook stream contains BIFF8 records
        guard let entry = cfb.root.children.first(where: { $0.name == "Workbook" }),
              let reader = try? cfb.stream(entry) else {
            throw DocReaderError.corruptedFile
        }

        return countBoundSheetRecords(in: reader.readData(ofLength: reader.totalBytes))
    }

    /// Scans BIFF8 records for BOUNDSHEET (type 0x0085) to count worksheets.
    private func countBoundSheetRecords(in data: Data) -> Int {
        var count = 0
        var offset = 0
        let bytes = [UInt8](data)

        while offset + 4 <= bytes.count {
            let recordType = UInt16(bytes[offset]) | (UInt16(bytes[offset + 1]) << 8)
            let recordLen  = UInt16(bytes[offset + 2]) | (UInt16(bytes[offset + 3]) << 8)
            if recordType == 0x0085 {
                count += 1
            }
            offset += 4 + Int(recordLen)
        }
        return max(1, count)
    }

    private func parseMetadata() throws -> DocumentMetadata {
        do {
            let cfb = try OLEFile(url.path)
            guard let entry = cfb.root.children.first(where: { $0.name == "\u{05}SummaryInformation" }),
                  let reader = try? cfb.stream(entry) else { return DocumentMetadata() }
            let data     = reader.readData(ofLength: reader.totalBytes)
            let title    = OLEPropertySetReader.readString(from: data, propertyID: 0x02)
            let author   = OLEPropertySetReader.readString(from: data, propertyID: 0x04)
            let created  = OLEPropertySetReader.readDate(from: data, propertyID: 0x0C)
            let modified = OLEPropertySetReader.readDate(from: data, propertyID: 0x0D)
            return DocumentMetadata(title: title, author: author,
                                    modifiedDate: modified, createdDate: created)
        } catch {
            return DocumentMetadata()
        }
    }
}
