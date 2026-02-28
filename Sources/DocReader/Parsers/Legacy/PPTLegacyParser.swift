import Foundation
import CoreGraphics
import OLEKit

/// Parses `.ppt` files (PowerPoint 97-2003 binary format via OLE2).
actor PPTLegacyParser: DocReadable {
    nonisolated let url: URL
    nonisolated let format: DocumentFormat = .ppt

    private var _pageCount: Int?
    private var _pageSize: CGSize?
    private var _metadata: DocumentMetadata?

    init(url: URL) {
        self.url = url
    }

    var pageCount: Int {
        get throws {
            if let cached = _pageCount { return cached }
            let count = try parseSlideCount()
            _pageCount = count
            return count
        }
    }

    func pageSize(at index: Int) throws -> CGSize {
        if let cached = _pageSize { return cached }
        let size = (try? parseSlideSize()) ?? CGSize(width: 720, height: 540)
        _pageSize = size
        return size
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

    private func parseSlideCount() throws -> Int {
        let data = try readPowerPointStream()
        return countSlideContainerAtoms(in: data)
    }

    private func parseSlideSize() throws -> CGSize {
        let data = try readPowerPointStream()
        return readDocumentAtomSize(from: data) ?? CGSize(width: 720, height: 540)
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

    // MARK: - OLE helpers

    private func readPowerPointStream() throws -> Data {
        let cfb: OLEFile
        do {
            cfb = try OLEFile(url.path)
        } catch {
            throw DocReaderError.corruptedFile
        }
        guard let entry = cfb.root.children.first(where: { $0.name == "PowerPoint Document" }),
              let reader = try? cfb.stream(entry) else {
            throw DocReaderError.corruptedFile
        }
        return reader.readData(ofLength: reader.totalBytes)
    }

    /// Counts SlideContainer records (type 0x03E8) in the PowerPoint stream.
    private func countSlideContainerAtoms(in data: Data) -> Int {
        var count = 0
        var offset = 0
        let bytes = [UInt8](data)

        while offset + 8 <= bytes.count {
            // PPT record header: version+instance (2), type (2), length (4)
            let recType = UInt16(bytes[offset + 2]) | (UInt16(bytes[offset + 3]) << 8)
            let recLen  = UInt32(bytes[offset + 4])
                        | (UInt32(bytes[offset + 5]) << 8)
                        | (UInt32(bytes[offset + 6]) << 16)
                        | (UInt32(bytes[offset + 7]) << 24)
            if recType == 0x03E8 {
                count += 1
            }
            offset += 8 + Int(recLen)
        }
        return max(1, count)
    }

    /// Reads DocumentAtom (type 0x03E9) to extract slide dimensions in EMU.
    /// Returns size in points.
    private func readDocumentAtomSize(from data: Data) -> CGSize? {
        var offset = 0
        let bytes = [UInt8](data)

        while offset + 8 <= bytes.count {
            let recType = UInt16(bytes[offset + 2]) | (UInt16(bytes[offset + 3]) << 8)
            let recLen  = UInt32(bytes[offset + 4])
                        | (UInt32(bytes[offset + 5]) << 8)
                        | (UInt32(bytes[offset + 6]) << 16)
                        | (UInt32(bytes[offset + 7]) << 24)

            if recType == 0x03E9 && recLen >= 8 && offset + 8 + 8 <= bytes.count {
                let bodyOffset = offset + 8
                let cx = UInt32(bytes[bodyOffset])
                       | (UInt32(bytes[bodyOffset + 1]) << 8)
                       | (UInt32(bytes[bodyOffset + 2]) << 16)
                       | (UInt32(bytes[bodyOffset + 3]) << 24)
                let cy = UInt32(bytes[bodyOffset + 4])
                       | (UInt32(bytes[bodyOffset + 5]) << 8)
                       | (UInt32(bytes[bodyOffset + 6]) << 16)
                       | (UInt32(bytes[bodyOffset + 7]) << 24)
                // PPT dimensions are in master units (1/576 inch) → points (1/72 inch)
                // master units to points: ÷ 8
                let width  = CGFloat(cx) / 8.0
                let height = CGFloat(cy) / 8.0
                if width > 0, height > 0 { return CGSize(width: width, height: height) }
            }
            offset += 8 + Int(recLen)
        }
        return nil
    }
}
