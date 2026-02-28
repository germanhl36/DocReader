/// The file format of a Microsoft Office document.
public enum DocumentFormat: String, Sendable, CaseIterable {
    // OOXML (modern, ZIP-based)
    case docx
    case xlsx
    case pptx

    // Legacy binary (CFB / OLE2)
    case doc
    case xls
    case ppt

    /// High-level grouping for a document format.
    public enum DocumentFamily: Sendable {
        case word
        case spreadsheet
        case presentation
    }

    /// The family this format belongs to.
    public var family: DocumentFamily {
        switch self {
        case .docx, .doc: return .word
        case .xlsx, .xls: return .spreadsheet
        case .pptx, .ppt: return .presentation
        }
    }

    /// Whether this is a legacy binary (CFB/OLE2) format.
    public var isLegacy: Bool {
        switch self {
        case .doc, .xls, .ppt: return true
        case .docx, .xlsx, .pptx: return false
        }
    }

    /// Whether this is a modern OOXML format.
    public var isOOXML: Bool { !isLegacy }

    /// File extension string (no dot).
    public var fileExtension: String { rawValue }
}
