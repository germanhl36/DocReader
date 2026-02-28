import Foundation
import ZIPFoundation

/// Wraps ZIPFoundation to extract entries from OOXML (ZIP) archives.
final class OOXMLZipExtractor: Sendable {
    let url: URL

    init(url: URL) {
        self.url = url
    }

    /// Extracts the ZIP entry at `entryPath` into memory.
    ///
    /// - Throws: ``DocReaderError/corruptedFile`` if the path does not exist inside the archive.
    func extractEntry(path entryPath: String) throws -> Data {
        guard let archive = Archive(url: url, accessMode: .read) else {
            throw DocReaderError.corruptedFile
        }
        guard let entry = archive[entryPath] else {
            throw DocReaderError.corruptedFile
        }
        var result = Data()
        do {
            _ = try archive.extract(entry) { chunk in
                result.append(chunk)
            }
        } catch {
            throw DocReaderError.corruptedFile
        }
        return result
    }

    /// Validates that this archive contains a `[Content_Types].xml` entry,
    /// which is required by the OOXML specification.
    ///
    /// - Throws: ``DocReaderError/corruptedFile`` if the entry is missing.
    func validateOOXMLStructure() throws {
        guard let archive = Archive(url: url, accessMode: .read) else {
            throw DocReaderError.corruptedFile
        }
        guard archive["[Content_Types].xml"] != nil else {
            throw DocReaderError.corruptedFile
        }
    }
}
