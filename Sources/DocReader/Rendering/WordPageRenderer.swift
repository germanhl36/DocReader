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
        var prevSpacingAfterPt: CGFloat = 0
        var prevBackgroundHex: String? = nil

        for element in content.elements {
            switch element {
            case .paragraph(let para):
                guard para.styleName != "__pagebreak__" else { continue }
                drawParagraph(para, in: context, contentRect: contentRect,
                              currentY: &currentY, prevSpacingAfterPt: &prevSpacingAfterPt,
                              prevBackgroundHex: prevBackgroundHex)
                prevBackgroundHex = para.backgroundHex
            case .table(let table):
                drawTable(table, in: context, contentRect: contentRect, currentY: &currentY)
                prevSpacingAfterPt = 0
                prevBackgroundHex = nil
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
            let colWidths = columnWidths(for: table, availableWidth: availableWidth)
            return table.rows.reduce(0) { $0 + rowHeight(for: $1, columnWidths: colWidths) } + 4
        }
    }

    /// Computes the rendered width of each column, scaled to fill availableWidth.
    static func columnWidths(for table: WordTableContent, availableWidth: CGFloat) -> [CGFloat] {
        let totalSpec = table.columnWidthsPt.reduce(0, +)
        if totalSpec > 0 && !table.columnWidthsPt.isEmpty {
            let scale = availableWidth / totalSpec
            return table.columnWidthsPt.map { $0 * scale }
        }
        let colCount = max(table.rows.map { $0.cells.count }.max() ?? 1, 1)
        let w = availableWidth / CGFloat(colCount)
        return Array(repeating: w, count: colCount)
    }

    /// Measures the height of a single table row based on its cell content and per-column widths.
    static func rowHeight(for row: WordTableRow, columnWidths: [CGFloat]) -> CGFloat {
        let minHeight: CGFloat = 20
        let maxCellHeight = row.cells.enumerated().compactMap { (ci, cell) -> CGFloat? in
            let cw = ci < columnWidths.count ? columnWidths[ci] : (columnWidths.last ?? minHeight)
            let hMargin: CGFloat = cell.margins.map { $0.left + $0.right } ?? 4.0
            let vMargin: CGFloat = cell.margins.map { $0.top + $0.bottom } ?? 4.0
            let textWidth = max(cw - hMargin, 1)
            var totalH: CGFloat = 0
            for para in cell.paragraphs {
                let attrStr = buildSingleParagraphAttrStr(para)
                let fs = CTFramesetterCreateWithAttributedString(attrStr)
                let fit = CTFramesetterSuggestFrameSizeWithConstraints(
                    fs, CFRangeMake(0, 0), nil,
                    CGSize(width: textWidth, height: .greatestFiniteMagnitude), nil)
                totalH += para.spacingBeforePt + (fit.height > 0 ? fit.height : 0) + para.spacingAfterPt
            }
            return totalH > 0 ? totalH + vMargin : nil
        }.max() ?? minHeight
        return max(maxCellHeight, minHeight)
    }

    /// Legacy overload used by tests — keeps a single uniform column width for one-column tables.
    static func rowHeight(for row: WordTableRow, colWidth: CGFloat) -> CGFloat {
        rowHeight(for: row, columnWidths: [colWidth])
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
        currentY: inout CGFloat,
        prevSpacingAfterPt: inout CGFloat,
        prevBackgroundHex: String?
    ) {
        // Paragraph spacing collapse (Word rule): the gap between two adjacent paragraphs
        // is max(prevAfter, thisBefore), not the sum. Only apply the additional spacing
        // beyond what the previous paragraph's spacing-after already consumed.
        let additionalBefore = max(0, para.spacingBeforePt - prevSpacingAfterPt)
        currentY -= additionalBefore

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
            prevSpacingAfterPt = para.spacingAfterPt
            return
        }

        // Draw shading background.
        // Extend into spacing-after so consecutive shaded paragraphs form a solid band.
        // When spacing was collapsed with a preceding same-background paragraph, also extend
        // upward by prevSpacingAfterPt to overdraw that paragraph's spacing-after zone —
        // this guarantees overlap and eliminates sub-pixel seam artifacts at shared edges.
        if let bgHex = para.backgroundHex, let bgColor = resolveColor(bgHex) {
            let topExtension: CGFloat = (additionalBefore == 0 && prevBackgroundHex == bgHex)
                ? prevSpacingAfterPt : 0
            let bgRect = CGRect(
                x: contentRect.minX,
                y: currentY - para.spacingAfterPt,
                width: contentRect.width,
                height: height + topExtension + para.spacingAfterPt
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

        // Draw paragraph top border (above text area)
        if let hex = para.borderTopHex, let borderColor = resolveColor(hex),
           para.borderTopWidthPt > 0 {
            context.saveGState()
            context.setStrokeColor(borderColor)
            context.setLineWidth(para.borderTopWidthPt)
            let y = currentY + height
            context.move(to: CGPoint(x: contentRect.minX, y: y))
            context.addLine(to: CGPoint(x: contentRect.maxX, y: y))
            context.strokePath()
            context.restoreGState()
        }

        // Draw paragraph bottom border (below text area)
        if let hex = para.borderBottomHex, let borderColor = resolveColor(hex),
           para.borderBottomWidthPt > 0 {
            context.saveGState()
            context.setStrokeColor(borderColor)
            context.setLineWidth(para.borderBottomWidthPt)
            context.move(to: CGPoint(x: contentRect.minX, y: currentY))
            context.addLine(to: CGPoint(x: contentRect.maxX, y: currentY))
            context.strokePath()
            context.restoreGState()
        }

        currentY -= para.spacingAfterPt
        prevSpacingAfterPt = para.spacingAfterPt
    }

    private static func drawTable(
        _ table: WordTableContent,
        in context: CGContext,
        contentRect: CGRect,
        currentY: inout CGFloat
    ) {
        guard !table.rows.isEmpty else { return }

        // Column widths (shared with measureElement for consistent layout)
        let colWidths = columnWidths(for: table, availableWidth: contentRect.width)

        let xForCol: (Int) -> CGFloat = { colIdx in
            contentRect.minX + colWidths.prefix(colIdx).reduce(0, +)
        }
        let widthForCol: (Int) -> CGFloat = { colIdx in
            colIdx < colWidths.count ? colWidths[colIdx] : (colWidths.last ?? contentRect.width)
        }

        // Compute per-row heights using actual cell widths
        let rowHeights: [CGFloat] = table.rows.map { rowHeight(for: $0, columnWidths: colWidths) }
        let totalHeight = rowHeights.reduce(0, +)

        currentY -= totalHeight
        guard currentY >= contentRect.minY - totalHeight else {
            currentY -= 4
            return
        }

        var rowY = currentY + totalHeight

        for (rowIdx, row) in table.rows.enumerated() {
            let rh = rowHeights[rowIdx]
            rowY -= rh

            for (colIdx, cell) in row.cells.enumerated() {
                let cx = xForCol(colIdx)
                let cw = widthForCol(colIdx)
                let cellRect = CGRect(x: cx, y: rowY, width: cw, height: rh)

                // Cell background — explicit color, then header default
                let bgColor: CGColor
                if let hex = cell.backgroundHex, let c = resolveColor(hex) {
                    bgColor = c
                } else if row.isHeader {
                    bgColor = CGColor(red: 0.25, green: 0.25, blue: 0.45, alpha: 1)
                } else {
                    bgColor = CGColor(red: 1, green: 1, blue: 1, alpha: 1)
                }
                context.setFillColor(bgColor)
                context.fill(cellRect)

                // Grid lines — use cell border color if available
                let borderStroke: CGColor
                if let hex = cell.borderColorHex, let c = resolveColor(hex) {
                    borderStroke = c
                } else {
                    borderStroke = CGColor(red: 0.5, green: 0.5, blue: 0.5, alpha: 1)
                }
                context.setStrokeColor(borderStroke)
                context.setLineWidth(0.5)
                context.stroke(cellRect)

                // Cell content — render each paragraph with proper formatting
                let hMarginL: CGFloat = cell.margins?.left   ?? 2
                let hMarginR: CGFloat = cell.margins?.right  ?? 2
                let vMarginT: CGFloat = cell.margins?.top    ?? 2
                let vMarginB: CGFloat = cell.margins?.bottom ?? 2
                let textRect = CGRect(
                    x: cellRect.minX + hMarginL,
                    y: cellRect.minY + vMarginB,
                    width: cellRect.width - hMarginL - hMarginR,
                    height: cellRect.height - vMarginT - vMarginB
                )
                var cellY = textRect.maxY
                for para in cell.paragraphs {
                    guard cellY > textRect.minY else { break }
                    cellY -= para.spacingBeforePt
                    let attrStr = buildSingleParagraphAttrStr(para)
                    let fs = CTFramesetterCreateWithAttributedString(attrStr)
                    let fit = CTFramesetterSuggestFrameSizeWithConstraints(
                        fs, CFRangeMake(0, 0), nil,
                        CGSize(width: max(textRect.width, 1), height: .greatestFiniteMagnitude), nil)
                    let ph = fit.height > 0 ? fit.height : 0
                    cellY -= ph
                    guard cellY >= textRect.minY - ph else { break }
                    let paraPath = CGPath(
                        rect: CGRect(x: textRect.minX, y: cellY, width: textRect.width, height: ph),
                        transform: nil)
                    let paraFrame = CTFramesetterCreateFrame(fs, CFRangeMake(0, 0), paraPath, nil)
                    CTFrameDraw(paraFrame, context)
                    cellY -= para.spacingAfterPt
                }
            }
        }

        currentY -= 4
    }

    // MARK: - Attributed string builders

    private static func buildSingleParagraphAttrStr(_ para: WordParagraphContent) -> CFAttributedString {
        let cfStr = CFAttributedStringCreateMutable(nil, 0)!
        var offset = 0

        // Prepend list prefix as a virtual run inheriting font from first text run
        var allRuns = para.runs
        if let prefix = para.listPrefix, !prefix.isEmpty {
            let firstRun = para.runs.first
            let prefixRun = WordRunContent(
                text: prefix, bold: false, italic: false,
                fontSizePt: firstRun?.fontSizePt ?? 0,
                hexColor: nil,
                fontFamily: firstRun?.fontFamily)
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
            if run.underline {
                let v = NSNumber(value: CTUnderlineStyle.single.rawValue)
                CFAttributedStringSetAttribute(cfStr, CFRangeMake(offset, len),
                                               kCTUnderlineStyleAttributeName, v)
            }
            if run.strikethrough {
                let v = NSNumber(value: CTUnderlineStyle.single.rawValue)
                CFAttributedStringSetAttribute(cfStr, CFRangeMake(offset, len),
                                               "NSStrikethrough" as CFString, v)
            }
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
            let defaultFont = CTFontCreateWithName("Arial" as CFString, 10, nil)
            CFAttributedStringBeginEditing(cfStr)
            CFAttributedStringReplaceString(cfStr, CFRangeMake(0, 0), nl)
            CFAttributedStringSetAttribute(cfStr, CFRangeMake(0, 1),
                                           kCTFontAttributeName, defaultFont)
            CFAttributedStringEndEditing(cfStr)
        }

        return cfStr
    }

    private static func buildCTParagraphStyle(_ para: WordParagraphContent) -> CTParagraphStyle {
        // Spacing (spacingBefore/After) is handled externally via currentY adjustments with
        // paragraph spacing collapse. Do NOT pass them into the CTParagraphStyle — doing so
        // would cause them to be included in CTFramesetterSuggestFrameSizeWithConstraints
        // and then double-counted when we also subtract them from currentY.
        var alignment = para.alignment
        var headIndent: CGFloat = para.leftIndentPt
        var firstLineIndent: CGFloat = para.leftIndentPt + para.firstLineIndentPt

        return withUnsafePointer(to: &alignment) { aPtr in
            withUnsafePointer(to: &headIndent) { hPtr in
                withUnsafePointer(to: &firstLineIndent) { fPtr in
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
                    ]
                    return CTParagraphStyleCreate(settings, settings.count)
                }
            }
        }
    }

    private static func resolveFont(run: WordRunContent, styleName: String) -> CTFont {
        let isHeading = styleName.lowercased().hasPrefix("heading") || styleName == "Title"
        let isBold = run.bold || isHeading
        let isItalic = run.italic
        let fontSize = run.fontSizePt > 0 ? run.fontSizePt : (isHeading ? 16 : 12)

        let base = CTFontCreateWithName(
            (run.fontFamily ?? "Helvetica") as CFString, fontSize, nil)

        if isBold || isItalic {
            var traits: CTFontSymbolicTraits = []
            if isBold   { traits.insert(.traitBold) }
            if isItalic { traits.insert(.traitItalic) }
            if let styled = CTFontCreateCopyWithSymbolicTraits(
                    base, fontSize, nil, traits, traits) {
                return styled
            }
        }
        return base
    }

    private static func isDark(_ color: CGColor) -> Bool {
        guard let comps = color.components, comps.count >= 3 else { return false }
        let luminance = 0.299 * comps[0] + 0.587 * comps[1] + 0.114 * comps[2]
        return luminance < 0.5
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
