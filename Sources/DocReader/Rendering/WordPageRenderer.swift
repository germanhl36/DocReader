import Foundation
import CoreGraphics
import CoreText

/// Renders Word document pages into a CoreGraphics context.
enum WordPageRenderer {

    // MARK: - Real content rendering

    @DocRenderActor
    static func render(in context: CGContext, content: WordPageContent) {
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
        context.fill(CGRect(origin: .zero, size: content.pageSize))

        let m = content.margins
        let contentRect = CGRect(
            x: m.left,
            y: m.bottom,
            width: content.pageSize.width - m.left - m.right,
            height: content.pageSize.height - m.top - m.bottom
        )

        let cfAttrStr = buildAttributedString(from: content.paragraphs)
        let framesetter = CTFramesetterCreateWithAttributedString(cfAttrStr)
        let path = CGPath(rect: contentRect, transform: nil)
        let frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path, nil)
        CTFrameDraw(frame, context)
    }

    // MARK: - Placeholder rendering

    @DocRenderActor
    static func renderPlaceholder(in context: CGContext, size: CGSize, pageIndex: Int) {
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
        context.fill(CGRect(origin: .zero, size: size))

        let margin: CGFloat = 72
        let contentRect = CGRect(
            x: margin, y: margin,
            width: size.width - margin * 2,
            height: size.height - margin * 2
        )
        context.setStrokeColor(CGColor(red: 0.9, green: 0.9, blue: 0.9, alpha: 1))
        context.setLineWidth(0.5)
        context.stroke(contentRect)

        renderCTText(
            "Page \(pageIndex + 1)",
            in: context,
            rect: CGRect(x: margin, y: size.height - margin - 40, width: contentRect.width, height: 40),
            fontSize: 14,
            color: CGColor(red: 0.3, green: 0.3, blue: 0.3, alpha: 1)
        )
    }

    // MARK: - Internal helpers (accessible for tests)

    /// Resolves a 6-character hex color string to a CGColor, or nil if invalid.
    static func resolveColor(_ hex: String) -> CGColor? {
        let clean = hex.trimmingCharacters(in: .whitespaces).replacingOccurrences(of: "#", with: "")
        guard clean.count == 6, let value = UInt64(clean, radix: 16) else { return nil }
        let r = CGFloat((value >> 16) & 0xFF) / 255.0
        let g = CGFloat((value >>  8) & 0xFF) / 255.0
        let b = CGFloat( value        & 0xFF) / 255.0
        return CGColor(red: r, green: g, blue: b, alpha: 1)
    }

    // MARK: - Private helpers

    private static func buildAttributedString(from paragraphs: [WordParagraphContent]) -> CFAttributedString {
        let cfStr = CFAttributedStringCreateMutable(nil, 0)!
        var offset = 0

        for para in paragraphs {
            guard para.styleName != "__pagebreak__" else { continue }

            for run in para.runs {
                let cfText = run.text as CFString
                let len = CFStringGetLength(cfText)
                guard len > 0 else { continue }

                let font = resolveFont(run: run, styleName: para.styleName)
                let color: CGColor = run.hexColor.flatMap { resolveColor($0) }
                    ?? CGColor(red: 0, green: 0, blue: 0, alpha: 1)

                CFAttributedStringBeginEditing(cfStr)
                CFAttributedStringReplaceString(cfStr, CFRangeMake(offset, 0), cfText)
                CFAttributedStringSetAttribute(cfStr, CFRangeMake(offset, len),
                                               kCTFontAttributeName, font)
                CFAttributedStringSetAttribute(cfStr, CFRangeMake(offset, len),
                                               kCTForegroundColorAttributeName, color)
                CFAttributedStringEndEditing(cfStr)
                offset += len
            }

            // Paragraph newline with default font
            let nl = "\n" as CFString
            let defaultFont = CTFontCreateWithName("Helvetica" as CFString, 12, nil)
            CFAttributedStringBeginEditing(cfStr)
            CFAttributedStringReplaceString(cfStr, CFRangeMake(offset, 0), nl)
            CFAttributedStringSetAttribute(cfStr, CFRangeMake(offset, 1),
                                           kCTFontAttributeName, defaultFont)
            CFAttributedStringEndEditing(cfStr)
            offset += 1
        }

        return cfStr
    }

    private static func resolveFont(run: WordRunContent, styleName: String) -> CTFont {
        let isHeading = styleName.lowercased().hasPrefix("heading") || styleName == "Title"
        let isBold = run.bold || isHeading
        let isItalic = run.italic
        let fontSize = run.fontSizePt > 0 ? run.fontSizePt : (isHeading ? 16 : 12)

        let name: CFString
        if isBold && isItalic {
            name = "Helvetica-BoldOblique" as CFString
        } else if isBold {
            name = "Helvetica-Bold" as CFString
        } else if isItalic {
            name = "Helvetica-Oblique" as CFString
        } else {
            name = "Helvetica" as CFString
        }
        return CTFontCreateWithName(name, fontSize, nil)
    }

    private static func renderCTText(
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
        let attrStr = CFAttributedStringCreate(nil, text as CFString, attributes as CFDictionary)!
        let framesetter = CTFramesetterCreateWithAttributedString(attrStr)
        let path = CGPath(rect: rect, transform: nil)
        let frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path, nil)
        CTFrameDraw(frame, context)
    }
}
