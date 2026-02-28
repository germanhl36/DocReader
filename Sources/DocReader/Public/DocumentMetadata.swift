import Foundation

/// Metadata extracted from a document's property streams.
public struct DocumentMetadata: Sendable, Equatable {
    /// Document title, if available.
    public let title: String?

    /// Author / creator, if available.
    public let author: String?

    /// Last modified date, if available.
    public let modifiedDate: Date?

    /// Creation date, if available.
    public let createdDate: Date?

    /// Application that created the document, if available.
    public let application: String?

    public init(
        title: String? = nil,
        author: String? = nil,
        modifiedDate: Date? = nil,
        createdDate: Date? = nil,
        application: String? = nil
    ) {
        self.title = title
        self.author = author
        self.modifiedDate = modifiedDate
        self.createdDate = createdDate
        self.application = application
    }
}
