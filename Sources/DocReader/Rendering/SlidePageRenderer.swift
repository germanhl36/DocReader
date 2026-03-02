import Foundation
import CoreGraphics
import CoreText

/// Renders presentation slides into a CoreGraphics context.
enum SlidePageRenderer {

    // MARK: - Real content rendering

    @DocRenderActor
    static func render(in context: CGContext, size: CGSize, content: SlidePageContent) {
        // Dark background
        context.setFillColor(CGColor(red: 0.18, green: 0.18, blue: 0.35, alpha: 1))
        context.fill(CGRect(origin: .zero, size: size))

        let white = CGColor(red: 1, green: 1, blue: 1, alpha: 1)

        for box in content.textBoxes {
            let text = box.lines.joined(separator: "\n")
            guard !text.isEmpty else { continue }

            let fontSize: CGFloat = box.isTitle ? 28 : 14
            let fontName = box.isTitle ? "Helvetica-Bold" : "Helvetica"

            // PPTX origin is top-left (Y-DOWN). Convert to PDF Y-UP.
            let pdfFrame = CGRect(
                x: box.frame.origin.x,
                y: size.height - box.frame.origin.y - box.frame.height,
                width: box.frame.width,
                height: box.frame.height
            )

            let attributes: [CFString: Any] = [
                kCTFontAttributeName: CTFontCreateWithName(fontName as CFString, fontSize, nil),
                kCTForegroundColorAttributeName: white
            ]
            let attrStr = CFAttributedStringCreate(nil, text as CFString, attributes as CFDictionary)!
            let framesetter = CTFramesetterCreateWithAttributedString(attrStr)
            let path = CGPath(rect: pdfFrame, transform: nil)
            let frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path, nil)
            CTFrameDraw(frame, context)
        }
    }

    // MARK: - Placeholder rendering

    @DocRenderActor
    static func renderPlaceholder(in context: CGContext, size: CGSize, pageIndex: Int) {
        let bgColor = CGColor(red: 0.18, green: 0.18, blue: 0.35, alpha: 1)
        context.setFillColor(bgColor)
        context.fill(CGRect(origin: .zero, size: size))

        // In Y-UP: y is measured from the bottom. Convert from Y-DOWN layout.
        let titleRect = CGRect(x: 60, y: size.height - 140, width: size.width - 120, height: 80)
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 0.08))
        context.fill(titleRect)

        let contentRect = CGRect(x: 60, y: 60, width: size.width - 120, height: size.height - 220)
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 0.04))
        context.fill(contentRect)

        renderCTText(
            "Slide \(pageIndex + 1)",
            in: context,
            rect: CGRect(x: 60, y: size.height - 120, width: 300, height: 50),
            fontSize: 28,
            color: CGColor(red: 1, green: 1, blue: 1, alpha: 1)
        )
    }

    // MARK: - Private helpers

    private static func renderCTText(
        _ text: String,
        in context: CGContext,
        rect: CGRect,
        fontSize: CGFloat,
        color: CGColor
    ) {
        let attributes: [CFString: Any] = [
            kCTFontAttributeName: CTFontCreateWithName("Helvetica-Bold" as CFString, fontSize, nil),
            kCTForegroundColorAttributeName: color
        ]
        let attrStr = CFAttributedStringCreate(nil, text as CFString, attributes as CFDictionary)!
        let framesetter = CTFramesetterCreateWithAttributedString(attrStr)
        let path = CGPath(rect: rect, transform: nil)
        let frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path, nil)
        CTFrameDraw(frame, context)
    }
}
