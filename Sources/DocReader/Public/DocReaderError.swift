import Foundation

/// Errors thrown by DocReader operations.
public enum DocReaderError: Error, Sendable {
    /// The file format is not supported.
    case unsupportedFormat
    /// The file could not be found at the given URL.
    case fileNotFound
    /// The file appears corrupted or is not a valid Office document.
    case corruptedFile
    /// The requested page index is outside the valid range.
    case pageOutOfRange
    /// The PDF export was cancelled before completion.
    case exportCancelled
    /// An unexpected internal error occurred.
    case internalError(String)
}

extension DocReaderError: LocalizedError {
    public var errorDescription: String? {
        let bundle = Bundle.module
        let table: String? = nil
        switch self {
        case .unsupportedFormat:
            return bundle.localizedString(
                forKey: "error.unsupportedFormat",
                value: "The file format is not supported.",
                table: table
            )
        case .fileNotFound:
            return bundle.localizedString(
                forKey: "error.fileNotFound",
                value: "The file could not be found.",
                table: table
            )
        case .corruptedFile:
            return bundle.localizedString(
                forKey: "error.corruptedFile",
                value: "The file appears to be corrupted or is not a valid Office document.",
                table: table
            )
        case .pageOutOfRange:
            return bundle.localizedString(
                forKey: "error.pageOutOfRange",
                value: "The requested page is outside the valid range.",
                table: table
            )
        case .exportCancelled:
            return bundle.localizedString(
                forKey: "error.exportCancelled",
                value: "The PDF export was cancelled.",
                table: table
            )
        case .internalError(let detail):
            let template = bundle.localizedString(
                forKey: "error.internalError",
                value: "An internal error occurred: %@",
                table: table
            )
            return String(format: template, detail)
        }
    }
}
