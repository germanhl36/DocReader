import Foundation

/// Entry point for opening and inspecting Microsoft Office documents.
///
/// ## Usage
/// ```swift
/// // Check support before opening
/// guard DocReader.isSupported(url: fileURL) else { return }
///
/// // Open the document
/// let doc = try await DocReader.open(url: fileURL)
/// let pages = try await doc.pageCount
/// let pdf = try await doc.exportPDF()
/// ```
public enum DocReader {
    /// Returns `true` if the file at `url` has a supported Office format extension.
    ///
    /// This is a fast, synchronous check based on file extension only.
    /// It does **not** read the file from disk.
    public static func isSupported(url: URL) -> Bool {
        DocFormatDetector.format(forExtension: url.pathExtension) != nil
    }

    /// Opens the document at `url` and returns a parser ready for inspection.
    ///
    /// - Throws: ``DocReaderError/fileNotFound`` if the file does not exist,
    ///   ``DocReaderError/unsupportedFormat`` if the extension is not supported.
    /// - Returns: A ``DocReadable`` conforming type appropriate for the format.
    public static func open(url: URL) async throws -> any DocReadable {
        guard FileManager.default.fileExists(atPath: url.path) else {
            throw DocReaderError.fileNotFound
        }

        guard let format = DocFormatDetector.detect(url: url) else {
            throw DocReaderError.unsupportedFormat
        }

        return try DocReaderFactory.makeParser(url: url, format: format)
    }
}
