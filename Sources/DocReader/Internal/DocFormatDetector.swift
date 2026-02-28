import Foundation

/// Detects the ``DocumentFormat`` of a file from its extension and magic bytes.
enum DocFormatDetector {
    // MARK: - Magic byte signatures

    /// ZIP local file header: `PK\x03\x04` — used by all OOXML formats.
    private static let zipMagic: [UInt8] = [0x50, 0x4B, 0x03, 0x04]

    /// OLE2 Compound File Binary header: `\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1`.
    private static let ole2Magic: [UInt8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]

    // MARK: - Detection

    /// Detects the document format from `url`, checking extension first,
    /// then falling back to reading magic bytes from disk.
    ///
    /// - Returns: The detected ``DocumentFormat``, or `nil` if unsupported.
    static func detect(url: URL) -> DocumentFormat? {
        // 1. Extension-based detection (fast path)
        if let format = format(forExtension: url.pathExtension) {
            return format
        }
        // 2. Magic-byte fallback (handles files without proper extensions)
        return detectByMagicBytes(url: url)
    }

    /// Detects the format based solely on the file extension (case-insensitive).
    static func format(forExtension ext: String) -> DocumentFormat? {
        switch ext.lowercased() {
        case "docx": return .docx
        case "xlsx": return .xlsx
        case "pptx": return .pptx
        case "doc":  return .doc
        case "xls":  return .xls
        case "ppt":  return .ppt
        default:     return nil
        }
    }

    // MARK: - Magic bytes

    private static func detectByMagicBytes(url: URL) -> DocumentFormat? {
        guard let handle = try? FileHandle(forReadingFrom: url) else { return nil }
        defer { try? handle.close() }

        let data = handle.readData(ofLength: 8)
        let bytes = [UInt8](data)

        if bytes.starts(with: zipMagic) {
            // ZIP → OOXML family; use extension to pick the exact type
            return format(forExtension: url.pathExtension)
        }

        if bytes.starts(with: ole2Magic) {
            // OLE2 → legacy family; use extension to pick the exact type
            // If extension is absent/wrong, default to .doc
            return format(forExtension: url.pathExtension) ?? .doc
        }

        return nil
    }
}
