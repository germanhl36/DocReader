import Foundation
import CoreGraphics
import CoreText

/// Renders a placeholder Word page into a CoreGraphics context.
///
/// In a full implementation this would parse paragraph runs from `word/document.xml`
/// and lay them out using `CTFramesetter`. Here we render a representative placeholder
/// with margin guides so that PDF output is valid and visually identifiable.
enum WordPageRenderer {
    @DocRenderActor
    static func render(in context: CGContext, size: CGSize, pageIndex: Int) {
        // White background
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
        context.fill(CGRect(origin: .zero, size: size))

        // Margin guides
        let margin: CGFloat = 72 // 1 inch
        let contentRect = CGRect(
            x: margin, y: margin,
            width: size.width - margin * 2,
            height: size.height - margin * 2
        )
        context.setStrokeColor(CGColor(red: 0.9, green: 0.9, blue: 0.9, alpha: 1))
        context.setLineWidth(0.5)
        context.stroke(contentRect)

        // Page label
        renderText(
            "Page \(pageIndex + 1)",
            in: context,
            rect: CGRect(x: margin, y: margin, width: contentRect.width, height: 40),
            fontSize: 14,
            color: CGColor(red: 0.3, green: 0.3, blue: 0.3, alpha: 1)
        )
    }

    private static func renderText(
        _ text: String,
        in context: CGContext,
        rect: CGRect,
        fontSize: CGFloat,
        color: CGColor
    ) {
        let attributes: [CFString: Any] = [
            kCTFontAttributeName: CTFontCreateWithName("Helvetica" as CFString, fontSize, nil),
            kCTForegroundColorAttributeName: color
        ]
        let attributedString = CFAttributedStringCreate(
            nil,
            text as CFString,
            attributes as CFDictionary
        )!
        let framesetter = CTFramesetterCreateWithAttributedString(attributedString)
        let path = CGPath(rect: rect, transform: nil)
        let frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path, nil)
        CTFrameDraw(frame, context)
    }
}
