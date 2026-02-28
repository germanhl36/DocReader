import Foundation
import CoreGraphics

/// A type that can read and inspect a Microsoft Office document.
///
/// Conforming types are responsible for parsing a specific file format
/// and providing document metadata, page dimensions, and PDF export.
public protocol DocReadable: Actor {
    /// The URL of the source document.
    var url: URL { get }

    /// The detected format of the document.
    var format: DocumentFormat { get }

    /// Total number of pages (or sheets / slides) in the document.
    var pageCount: Int { get async throws }

    /// Dimensions of the page at the given zero-based index, in points.
    func pageSize(at index: Int) async throws -> CGSize

    /// Metadata extracted from the document's property streams.
    var metadata: DocumentMetadata { get async throws }

    /// Exports the full document as PDF data.
    func exportPDF() async throws -> Data

    /// Exports a subset of pages as PDF data.
    ///
    /// - Parameter pages: A closed range of zero-based page indices.
    func exportPDF(pages: ClosedRange<Int>) async throws -> Data
}
