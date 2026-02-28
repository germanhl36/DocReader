import Foundation
import CoreGraphics
import CoreText

/// Renders a placeholder presentation slide into a CoreGraphics context.
///
/// A full implementation would render slide backgrounds and text boxes
/// extracted from the OOXML/PPT binary stream.
enum SlidePageRenderer {
    @DocRenderActor
    static func render(in context: CGContext, size: CGSize, pageIndex: Int) {
        // Gradient-style background (solid approximation)
        let bgColor = CGColor(red: 0.18, green: 0.18, blue: 0.35, alpha: 1)
        context.setFillColor(bgColor)
        context.fill(CGRect(origin: .zero, size: size))

        // Title area placeholder
        let titleRect = CGRect(x: 60, y: 60, width: size.width - 120, height: 80)
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 0.08))
        context.fill(titleRect)

        // Content area placeholder
        let contentRect = CGRect(x: 60, y: 160, width: size.width - 120, height: size.height - 220)
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 0.04))
        context.fill(contentRect)

        // Slide number label
        renderText(
            "Slide \(pageIndex + 1)",
            in: context,
            rect: CGRect(x: 60, y: 70, width: 300, height: 50),
            fontSize: 28,
            color: CGColor(red: 1, green: 1, blue: 1, alpha: 1)
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
