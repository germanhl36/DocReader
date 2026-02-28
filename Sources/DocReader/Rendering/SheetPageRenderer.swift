import Foundation
import CoreGraphics
import CoreText

/// Renders a placeholder spreadsheet page into a CoreGraphics context.
///
/// A full implementation would render cell grids from the parsed worksheet data.
enum SheetPageRenderer {
    @DocRenderActor
    static func render(in context: CGContext, size: CGSize, pageIndex: Int) {
        // White background
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
        context.fill(CGRect(origin: .zero, size: size))

        let margin: CGFloat = 36
        let gridOrigin = CGPoint(x: margin, y: margin)
        let cellWidth: CGFloat = 80
        let cellHeight: CGFloat = 20
        let cols = Int((size.width - margin * 2) / cellWidth)
        let rows = Int((size.height - margin * 2) / cellHeight)

        context.setStrokeColor(CGColor(red: 0.85, green: 0.85, blue: 0.85, alpha: 1))
        context.setLineWidth(0.5)

        // Draw grid lines
        for col in 0...cols {
            let x = gridOrigin.x + CGFloat(col) * cellWidth
            context.move(to: CGPoint(x: x, y: gridOrigin.y))
            context.addLine(to: CGPoint(x: x, y: gridOrigin.y + CGFloat(rows) * cellHeight))
        }
        for row in 0...rows {
            let y = gridOrigin.y + CGFloat(row) * cellHeight
            context.move(to: CGPoint(x: gridOrigin.x, y: y))
            context.addLine(to: CGPoint(x: gridOrigin.x + CGFloat(cols) * cellWidth, y: y))
        }
        context.strokePath()

        // Sheet label
        renderText(
            "Sheet \(pageIndex + 1)",
            in: context,
            rect: CGRect(x: margin, y: margin, width: 200, height: cellHeight),
            fontSize: 10
        )
    }

    private static func renderText(
        _ text: String,
        in context: CGContext,
        rect: CGRect,
        fontSize: CGFloat
    ) {
        let color = CGColor(red: 0.2, green: 0.2, blue: 0.2, alpha: 1)
        let attributes: [CFString: Any] = [
            kCTFontAttributeName: CTFontCreateWithName("Helvetica" as CFString, fontSize, nil),
            kCTForegroundColorAttributeName: color
        ]
        let attrStr = CFAttributedStringCreate(nil, text as CFString, attributes as CFDictionary)!
        let framesetter = CTFramesetterCreateWithAttributedString(attrStr)
        let path = CGPath(rect: rect, transform: nil)
        let frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path, nil)
        CTFrameDraw(frame, context)
    }
}
