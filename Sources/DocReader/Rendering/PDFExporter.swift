import Foundation
import CoreGraphics

/// Produces PDF data from a ``DocReadable`` document using CoreGraphics.
///
/// All work is isolated to ``DocRenderActor`` to keep the main thread free.
enum PDFExporter {
    /// Exports `pages` from `parser` to in-memory PDF data.
    ///
    /// - Throws: ``DocReaderError/pageOutOfRange`` if any index is invalid,
    ///   ``DocReaderError/exportCancelled`` if the `Task` is cancelled.
    @DocRenderActor
    static func export(parser: any DocReadable, pages: ClosedRange<Int>) async throws -> Data {
        let pageCount = try await parser.pageCount

        guard pages.lowerBound >= 0, pages.upperBound < pageCount else {
            throw DocReaderError.pageOutOfRange
        }

        // Determine the format family for renderer selection
        let format = await parser.format

        // Use the first page's size as the default page rect
        let firstSize = try await parser.pageSize(at: pages.lowerBound)
        var mediaBox = CGRect(origin: .zero, size: firstSize)

        let pdfData = NSMutableData()
        guard let consumer = CGDataConsumer(data: pdfData),
              let context = CGContext(consumer: consumer, mediaBox: &mediaBox, nil) else {
            throw DocReaderError.internalError("Failed to create PDF context")
        }

        for pageIndex in pages {
            try Task.checkCancellation()

            let size = try await parser.pageSize(at: pageIndex)
            var pageRect = CGRect(origin: .zero, size: size)
            let boxData = NSData(bytes: &pageRect, length: MemoryLayout<CGRect>.size)
            let pageInfo: [CFString: Any] = [kCGPDFContextMediaBox: boxData]
            context.beginPDFPage(pageInfo as CFDictionary)

            // Flip coordinate system (CG origin is bottom-left)
            context.translateBy(x: 0, y: size.height)
            context.scaleBy(x: 1, y: -1)

            switch format.family {
            case .word:
                WordPageRenderer.render(in: context, size: size, pageIndex: pageIndex)
            case .spreadsheet:
                SheetPageRenderer.render(in: context, size: size, pageIndex: pageIndex)
            case .presentation:
                SlidePageRenderer.render(in: context, size: size, pageIndex: pageIndex)
            }

            context.endPDFPage()
        }

        context.closePDF()
        return pdfData as Data
    }
}
