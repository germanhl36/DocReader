import Foundation
import CoreGraphics
import CoreText

/// Renders Word document pages into a CoreGraphics context using per-element vertical layout.
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

        // Y-UP: start at the top of the content area
        var currentY = contentRect.maxY

        for element in content.elements {
            switch element {
            case .paragraph(let para):
                guard para.styleName != "__pagebreak__" else { continue }
                drawParagraph(para, in: context, contentRect: contentRect, currentY: &currentY)
            case .table(let table):
                drawTable(table, in: context, contentRect: contentRect, currentY: &currentY)
            }
        }
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

    // MARK: - Height measurement (used by splitIntoWordPages for page reflow)

    /// Returns the vertical height an element will occupy on a page, for layout purposes.
    static func measureElement(_ element: WordElement, availableWidth: CGFloat) -> CGFloat {
        switch element {
        case .paragraph(let para):
            guard para.styleName != "__pagebreak__" else { return 0 }
            let attrStr = buildSingleParagraphAttrStr(para)
            let framesetter = CTFramesetterCreateWithAttributedString(attrStr)
            let availW = max(availableWidth - para.leftIndentPt, 1)
            let maxSize = CGSize(width: availW, height: .greatestFiniteMagnitude)
            let fitSize = CTFramesetterSuggestFrameSizeWithConstraints(
                framesetter, CFRangeMake(0, 0), nil, maxSize, nil)
            return para.spacingBeforePt + max(fitSize.height, 14) + para.spacingAfterPt
        case .table(let table):
            guard !table.rows.isEmpty else { return 0 }
            let colCount = table.rows.map { $0.cells.count }.max() ?? 1
            let colWidth = colCount > 0 ? availableWidth / CGFloat(colCount) : availableWidth
            return table.rows.reduce(0) { $0 + rowHeight(for: $1, colWidth: colWidth) } + 4
        }
    }

    /// Measures the height of a single table row based on its cell content.
    static func rowHeight(for row: WordTableRow, colWidth: CGFloat) -> CGFloat {
        let minHeight: CGFloat = 20
        let maxCellHeight = row.cells.compactMap { cell -> CGFloat? in
            let text = cell.paragraphs.flatMap { $0.runs }.map { $0.text }.joined(separator: " ")
            guard !text.isEmpty else { return nil }
            let font = CTFontCreateWithName("Helvetica" as CFString, 9, nil)
            let attrs: [CFString: Any] = [kCTFontAttributeName: font]
            let attrStr = CFAttributedStringCreate(nil, text as CFString, attrs as CFDictionary)!
            let framesetter = CTFramesetterCreateWithAttributedString(attrStr)
            let maxSize = CGSize(width: max(colWidth - 4, 1), height: .greatestFiniteMagnitude)
            let fit = CTFramesetterSuggestFrameSizeWithConstraints(
                framesetter, CFRangeMake(0, 0), nil, maxSize, nil)
            return fit.height + 4
        }.max() ?? minHeight
        return max(maxCellHeight, minHeight)
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

    // MARK: - Per-element drawing

    private static func drawParagraph(
        _ para: WordParagraphContent,
        in context: CGContext,
        contentRect: CGRect,
        currentY: inout CGFloat
    ) {
        currentY -= para.spacingBeforePt

        let attrStr = buildSingleParagraphAttrStr(para)
        let framesetter = CTFramesetterCreateWithAttributedString(attrStr)

        let availWidth = contentRect.width - para.leftIndentPt
        let maxSize = CGSize(width: availWidth, height: CGFloat.greatestFiniteMagnitude)
        let fitSize = CTFramesetterSuggestFrameSizeWithConstraints(
            framesetter, CFRangeMake(0, 0), nil, maxSize, nil)

        let height = max(fitSize.height, 14)
        currentY -= height

        // Clip elements that overflow below the page
        guard currentY >= contentRect.minY - height else {
            currentY -= para.spacingAfterPt
            return
        }

        // Draw shading background
        if let bgHex = para.backgroundHex, let bgColor = resolveColor(bgHex) {
            let bgRect = CGRect(
                x: contentRect.minX,
                y: currentY,
                width: contentRect.width,
                height: height
            )
            context.setFillColor(bgColor)
            context.fill(bgRect)
        }

        // Draw paragraph text
        let textRect = CGRect(
            x: contentRect.minX + para.leftIndentPt,
            y: currentY,
            width: availWidth,
            height: height
        )
        let path = CGPath(rect: textRect, transform: nil)
        let frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path, nil)
        CTFrameDraw(frame, context)

        currentY -= para.spacingAfterPt
    }

    private static func drawTable(
        _ table: WordTableContent,
        in context: CGContext,
        contentRect: CGRect,
        currentY: inout CGFloat
    ) {
        guard !table.rows.isEmpty else { return }
        let colCount = table.rows.map { $0.cells.count }.max() ?? 1
        let colWidth = contentRect.width / CGFloat(colCount)

        // Compute per-row heights dynamically so cell text is never truncated
        let rowHeights = table.rows.map { rowHeight(for: $0, colWidth: colWidth) }
        let totalHeight = rowHeights.reduce(0, +)

        currentY -= totalHeight
        guard currentY >= contentRect.minY - totalHeight else {
            currentY -= 4
            return
        }

        var rowY = currentY + totalHeight  // Start at top of table

        for (rowIdx, row) in table.rows.enumerated() {
            let rh = rowHeights[rowIdx]
            rowY -= rh

            // Header row background
            if row.isHeader {
                let rowRect = CGRect(
                    x: contentRect.minX, y: rowY,
                    width: contentRect.width, height: rh)
                context.setFillColor(CGColor(red: 0.25, green: 0.25, blue: 0.45, alpha: 1))
                context.fill(rowRect)
            }

            for (colIdx, cell) in row.cells.enumerated() {
                let cellRect = CGRect(
                    x: contentRect.minX + CGFloat(colIdx) * colWidth,
                    y: rowY,
                    width: colWidth,
                    height: rh
                )

                // Grid lines
                context.setStrokeColor(CGColor(red: 0.5, green: 0.5, blue: 0.5, alpha: 1))
                context.setLineWidth(0.5)
                context.stroke(cellRect)

                // Cell text
                let text = cell.paragraphs.flatMap { $0.runs }.map { $0.text }.joined(separator: " ")
                guard !text.isEmpty else { continue }
                let textColor: CGColor = row.isHeader
                    ? CGColor(red: 1, green: 1, blue: 1, alpha: 1)
                    : CGColor(red: 0, green: 0, blue: 0, alpha: 1)
                renderCTText(text, in: context,
                             rect: cellRect.insetBy(dx: 2, dy: 2),
                             fontSize: 9, color: textColor)
            }
        }

        currentY -= 4  // Bottom spacing after table
    }

    // MARK: - Attributed string builders

    private static func buildSingleParagraphAttrStr(_ para: WordParagraphContent) -> CFAttributedString {
        let cfStr = CFAttributedStringCreateMutable(nil, 0)!
        var offset = 0

        // Prepend list prefix as a virtual run
        var allRuns = para.runs
        if let prefix = para.listPrefix, !prefix.isEmpty {
            let prefixRun = WordRunContent(
                text: prefix, bold: false, italic: false, fontSizePt: 0, hexColor: nil)
            allRuns = [prefixRun] + allRuns
        }

        for run in allRuns {
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

        // Apply paragraph style (alignment, indentation, spacing)
        if offset > 0 {
            let paraStyle = buildCTParagraphStyle(para)
            CFAttributedStringBeginEditing(cfStr)
            CFAttributedStringSetAttribute(cfStr, CFRangeMake(0, offset),
                                           kCTParagraphStyleAttributeName, paraStyle)
            CFAttributedStringEndEditing(cfStr)
        } else {
            // Empty paragraph — add a newline so it occupies space
            let nl = "\n" as CFString
            let defaultFont = CTFontCreateWithName("Helvetica" as CFString, 12, nil)
            CFAttributedStringBeginEditing(cfStr)
            CFAttributedStringReplaceString(cfStr, CFRangeMake(0, 0), nl)
            CFAttributedStringSetAttribute(cfStr, CFRangeMake(0, 1),
                                           kCTFontAttributeName, defaultFont)
            CFAttributedStringEndEditing(cfStr)
        }

        return cfStr
    }

    private static func buildCTParagraphStyle(_ para: WordParagraphContent) -> CTParagraphStyle {
        var alignment = para.alignment
        var headIndent: CGFloat = para.leftIndentPt
        var firstLineIndent: CGFloat = para.leftIndentPt + para.firstLineIndentPt
        var paragraphSpacing: CGFloat = para.spacingAfterPt
        var paragraphSpacingBefore: CGFloat = para.spacingBeforePt

        return withUnsafePointer(to: &alignment) { aPtr in
            withUnsafePointer(to: &headIndent) { hPtr in
                withUnsafePointer(to: &firstLineIndent) { fPtr in
                    withUnsafePointer(to: &paragraphSpacing) { spPtr in
                        withUnsafePointer(to: &paragraphSpacingBefore) { sbPtr in
                            let settings: [CTParagraphStyleSetting] = [
                                CTParagraphStyleSetting(
                                    spec: .alignment,
                                    valueSize: MemoryLayout<CTTextAlignment>.size,
                                    value: UnsafeRawPointer(aPtr)),
                                CTParagraphStyleSetting(
                                    spec: .headIndent,
                                    valueSize: MemoryLayout<CGFloat>.size,
                                    value: UnsafeRawPointer(hPtr)),
                                CTParagraphStyleSetting(
                                    spec: .firstLineHeadIndent,
                                    valueSize: MemoryLayout<CGFloat>.size,
                                    value: UnsafeRawPointer(fPtr)),
                                CTParagraphStyleSetting(
                                    spec: .paragraphSpacing,
                                    valueSize: MemoryLayout<CGFloat>.size,
                                    value: UnsafeRawPointer(spPtr)),
                                CTParagraphStyleSetting(
                                    spec: .paragraphSpacingBefore,
                                    valueSize: MemoryLayout<CGFloat>.size,
                                    value: UnsafeRawPointer(sbPtr)),
                            ]
                            return CTParagraphStyleCreate(settings, settings.count)
                        }
                    }
                }
            }
        }
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
