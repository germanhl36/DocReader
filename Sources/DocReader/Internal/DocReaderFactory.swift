import Foundation

/// Creates the appropriate ``DocReadable`` parser for a given URL and format.
enum DocReaderFactory {
    /// Returns a parser for `url`, or throws ``DocReaderError/unsupportedFormat``.
    ///
    /// Iteration 1: All parsers return stubs that throw `.unsupportedFormat`.
    static func makeParser(url: URL, format: DocumentFormat) throws -> any DocReadable {
        switch format {
        case .docx:
            return OOXMLWordParser(url: url)
        case .xlsx:
            return OOXMLSheetParser(url: url)
        case .pptx:
            return OOXMLSlideParser(url: url)
        case .doc:
            return DOCLegacyParser(url: url)
        case .xls:
            return XLSLegacyParser(url: url)
        case .ppt:
            return PPTLegacyParser(url: url)
        }
    }
}
