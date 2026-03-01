import Foundation
import CoreGraphics
import CoreText

/// Renders spreadsheet pages into a CoreGraphics context.
enum SheetPageRenderer {

    // MARK: - Real content rendering

    @DocRenderActor
    static func render(in context: CGContext, size: CGSize, content: SheetPageContent) {
        context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
        context.fill(CGRect(origin: .zero, size: size))

        let margin: CGFloat = 36
        let headerHeight: CGFloat = 24

        // Sheet name header
        renderCTText(
            content.sheetName,
            in: context,
            rect: CGRect(x: margin, y: margin, width: size.width - margin * 2, height: headerHeight),
            fontSize: 14,
            bold: true
        )

        guard !content.cells.isEmpty else { return }

        let maxCol = content.cells.map { $0.col }.max() ?? 0
        let maxRow = content.cells.map { $0.row }.max() ?? 0

        let rowLabelWidth: CGFloat = 30
        let gridOriginX = margin + rowLabelWidth
        let gridOriginY = margin + headerHeight + 8

        let availableWidth  = size.width  - gridOriginX - margin
        let availableHeight = size.height - gridOriginY - margin

        let colCount = max(1, maxCol + 1)
        let rowCount = max(1, maxRow + 1)
        let cellWidth  = min(80, availableWidth  / CGFloat(colCount))
        let cellHeight = min(20, availableHeight / CGFloat(rowCount))

        let visibleCols = Int(availableWidth  / cellWidth)
        let visibleRows = Int(availableHeight / cellHeight)

        // Column header row
        for col in 0...min(maxCol, visibleCols) {
            let x = gridOriginX + CGFloat(col) * cellWidth
            renderCTText(
                columnLabel(col),
                in: context,
                rect: CGRect(x: x + 1, y: gridOriginY, width: cellWidth - 2, height: cellHeight),
                fontSize: 8,
                bold: false
            )
        }

        // Row numbers
        for row in 0...min(maxRow, visibleRows) {
            let y = gridOriginY + CGFloat(row + 1) * cellHeight
            renderCTText(
                "\(row + 1)",
                in: context,
                rect: CGRect(x: margin, y: y, width: rowLabelWidth - 2, height: cellHeight),
                fontSize: 8,
                bold: false
            )
        }

        // Cell values
        for cell in content.cells {
            guard cell.col <= visibleCols, cell.row <= visibleRows else { continue }
            let x = gridOriginX + CGFloat(cell.col) * cellWidth
            let y = gridOriginY + CGFloat(cell.row + 1) * cellHeight
            renderCTText(
                cell.text,
                in: context,
                rect: CGRect(x: x + 2, y: y + 1, width: cellWidth - 4, height: cellHeight - 2),
                fontSize: 8,
                bold: false
            )
        }

        // Grid lines
        context.setStrokeColor(CGColor(red: 0.8, green: 0.8, blue: 0.8, alpha: 1))
        context.setLineWidth(0.5)

        let gridCols = min(maxCol, visibleCols) + 1
        let gridRows = min(maxRow, visibleRows) + 2

        for col in 0...gridCols {
            let x = gridOriginX + CGFloat(col) * cellWidth
            context.move(to: CGPoint(x: x, y: gridOriginY))
            context.addLine(to: CGPoint(x: x, y: gridOriginY + CGFloat(gridRows) * cellHeight))
        }
        for row in 0...gridRows {
            let y = gridOriginY + CGFloat(row) * cellHeight
            context.move(to: CGPoint(x: gridOriginX, y: y))
            context.addLine(to: CGPoint(x: gridOriginX + CGFloat(gridCols) * cellWidth, y: y))
        }
        context.strokePath()
    }

    // MARK: - Placeholder rendering

    @DocRenderActor
    static func renderPlaceholder(in context: CGContext, size: CGSize, pageIndex: Int) {
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

        renderCTText(
            "Sheet \(pageIndex + 1)",
            in: context,
            rect: CGRect(x: margin, y: margin, width: 200, height: cellHeight),
            fontSize: 10,
            bold: false
        )
    }

    // MARK: - Private helpers

    private static func renderCTText(
        _ text: String,
        in context: CGContext,
        rect: CGRect,
        fontSize: CGFloat,
        bold: Bool
    ) {
        let fontName = bold ? "Helvetica-Bold" : "Helvetica"
        let color = CGColor(red: 0.1, green: 0.1, blue: 0.1, alpha: 1)
        let attributes: [CFString: Any] = [
            kCTFontAttributeName: CTFontCreateWithName(fontName as CFString, fontSize, nil),
            kCTForegroundColorAttributeName: color
        ]
        let attrStr = CFAttributedStringCreate(nil, text as CFString, attributes as CFDictionary)!
        let framesetter = CTFramesetterCreateWithAttributedString(attrStr)
        let path = CGPath(rect: rect, transform: nil)
        let frame = CTFramesetterCreateFrame(framesetter, CFRangeMake(0, 0), path, nil)
        CTFrameDraw(frame, context)
    }
}
