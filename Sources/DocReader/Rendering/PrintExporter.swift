import Foundation
import CoreGraphics
import PDFKit

/// Rendered bitmap data for a single page, ready for printer-format encoding.
public struct PrintRasterPage: Sendable {
    /// Raw RGB (top-to-bottom) pixel data: width × height × 3 bytes.
    public let rgbData: Data
    public let widthPx: Int
    public let heightPx: Int
    /// Original PDF page width in points.
    public let widthPt: CGFloat
    /// Original PDF page height in points.
    public let heightPt: CGFloat
}

/// Converts PDF data into printer-native formats (PWG-Raster, URF, PCL 5, PCL XL).
///
/// All methods are isolated to ``DocRenderActor`` and can be called on any PDF `Data`.
public enum PrintExporter {

    // MARK: - PDF → Bitmap

    /// Renders every page of `pdf` into ``PrintRasterPage`` values at `resolution` DPI.
    @DocRenderActor
    public static func renderPages(pdf: Data, resolution: Int) throws -> [PrintRasterPage] {
        guard let document = PDFDocument(data: pdf) else {
            throw DocReaderError.corruptedFile
        }
        let scale = CGFloat(resolution) / 72.0
        var pages: [PrintRasterPage] = []
        for i in 0..<document.pageCount {
            guard let page = document.page(at: i) else { continue }
            let mediaBox = page.bounds(for: .mediaBox)
            let widthPx  = Int(ceil(mediaBox.width  * scale))
            let heightPx = Int(ceil(mediaBox.height * scale))

            // Create RGBA8888 context
            let bytesPerRow = widthPx * 4
            var rawBytes = [UInt8](repeating: 255, count: bytesPerRow * heightPx)
            guard let ctx = CGContext(
                data: &rawBytes,
                width: widthPx,
                height: heightPx,
                bitsPerComponent: 8,
                bytesPerRow: bytesPerRow,
                space: CGColorSpaceCreateDeviceRGB(),
                bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
            ) else {
                throw DocReaderError.internalError("CGContext creation failed")
            }

            // White background
            ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
            ctx.fill(CGRect(x: 0, y: 0, width: widthPx, height: heightPx))

            ctx.scaleBy(x: scale, y: scale)
            page.draw(with: .mediaBox, to: ctx)

            // Strip alpha: RGBA → RGB, also flip rows (CGContext is bottom-up)
            var rgb = [UInt8](repeating: 0, count: widthPx * heightPx * 3)
            for row in 0..<heightPx {
                let srcRow = heightPx - 1 - row   // flip vertical
                for col in 0..<widthPx {
                    let srcBase = srcRow * bytesPerRow + col * 4
                    let dstBase = row   * widthPx * 3 + col * 3
                    rgb[dstBase]     = rawBytes[srcBase]
                    rgb[dstBase + 1] = rawBytes[srcBase + 1]
                    rgb[dstBase + 2] = rawBytes[srcBase + 2]
                }
            }

            pages.append(PrintRasterPage(
                rgbData: Data(rgb),
                widthPx: widthPx,
                heightPx: heightPx,
                widthPt: mediaBox.width,
                heightPt: mediaBox.height
            ))
        }
        return pages
    }

    // MARK: - PWG-Raster

    /// Exports `pdf` as PWG-Raster (cups_page_header2_t, uncompressed RGB).
    @DocRenderActor
    public static func exportPWGRaster(pdf: Data, resolution: Int = 300) throws -> Data {
        let pages = try renderPages(pdf: pdf, resolution: resolution)
        var out = [UInt8]()
        // Magic "RaS2"
        out += Array("RaS2".utf8)

        for page in pages {
            // 1796-byte cups_page_header2_t, all zeros by default
            var hdr = [UInt8](repeating: 0, count: 1796)
            let res = UInt32(resolution)
            writeBEUInt32(res,                  into: &hdr, at: 276)  // HWResolutionX
            writeBEUInt32(res,                  into: &hdr, at: 280)  // HWResolutionY
            writeBEUInt32(0,                    into: &hdr, at: 284)  // ImagingBoundingBox left
            writeBEUInt32(0,                    into: &hdr, at: 288)  // ImagingBoundingBox bottom
            writeBEUInt32(UInt32(page.widthPx), into: &hdr, at: 292)  // ImagingBoundingBox right
            writeBEUInt32(UInt32(page.heightPx),into: &hdr, at: 296)  // ImagingBoundingBox top
            writeBEUInt32(1,                    into: &hdr, at: 340)  // NumCopies
            writeBEUInt32(UInt32(page.widthPt.rounded()), into: &hdr, at: 352)  // PageSizeX
            writeBEUInt32(UInt32(page.heightPt.rounded()),into: &hdr, at: 356)  // PageSizeY
            writeBEUInt32(UInt32(page.widthPx), into: &hdr, at: 372)  // cupsWidth
            writeBEUInt32(UInt32(page.heightPx),into: &hdr, at: 376)  // cupsHeight
            writeBEUInt32(8,                    into: &hdr, at: 384)  // cupsBitsPerColor
            writeBEUInt32(24,                   into: &hdr, at: 388)  // cupsBitsPerPixel
            writeBEUInt32(UInt32(page.widthPx * 3), into: &hdr, at: 392)  // cupsBytesPerLine
            writeBEUInt32(0,                    into: &hdr, at: 396)  // cupsColorOrder (chunky)
            writeBEUInt32(19,                   into: &hdr, at: 400)  // cupsColorSpace (sRGB)
            writeBEUInt32(0,                    into: &hdr, at: 404)  // cupsCompression (none)
            writeBEUInt32(3,                    into: &hdr, at: 420)  // cupsNumColors

            // cupsPageSize[0] and [1] as Float32 big-endian
            writeFloat32BE(Float(page.widthPt),  into: &hdr, at: 428)
            writeFloat32BE(Float(page.heightPt), into: &hdr, at: 432)

            out += hdr
            out += page.rgbData
        }
        return Data(out)
    }

    // MARK: - URF

    /// Exports `pdf` as Apple URF (UNIRAST) with PackBits compression.
    @DocRenderActor
    public static func exportURF(pdf: Data, resolution: Int = 300) throws -> Data {
        let pages = try renderPages(pdf: pdf, resolution: resolution)
        var out = [UInt8]()
        // Magic "UNIRAST\0"
        out += Array("UNIRAST\0".utf8)
        // Page count uint32 BE
        writeBEUInt32Into(&out, UInt32(pages.count))

        for page in pages {
            // 32-byte per-page header
            var hdr = [UInt8](repeating: 0, count: 32)
            hdr[0] = 24   // bitsPerPixel
            hdr[1] = 1    // colorspace sRGB
            hdr[2] = 0    // duplex off
            hdr[3] = 1    // quality normal
            writeBEUInt32(UInt32(page.widthPx),  into: &hdr, at: 12)
            writeBEUInt32(UInt32(page.heightPx), into: &hdr, at: 16)
            writeBEUInt32(UInt32(resolution),    into: &hdr, at: 20)
            out += hdr

            // PackBits compress each scanline
            let rowBytes = page.widthPx * 3
            let rgb = [UInt8](page.rgbData)
            for row in 0..<page.heightPx {
                let start = row * rowBytes
                let slice = rgb[start..<(start + rowBytes)]
                out += packBitsRow(slice)
            }
        }
        return Data(out)
    }

    // MARK: - PCL 5

    /// Exports `pdf` as PCL 5 raster.
    @DocRenderActor
    public static func exportPCL(pdf: Data, resolution: Int = 300) throws -> Data {
        let pages = try renderPages(pdf: pdf, resolution: resolution)
        var out = [UInt8]()
        out += [0x1B, 0x45]  // ESC E reset

        for page in pages {
            let paperCode = pclPaperCode(widthPt: page.widthPt)
            let rowBytes  = page.widthPx * 3
            let rgb       = [UInt8](page.rgbData)

            // Paper size
            out += pclEscStr("&l\(paperCode)A")
            // Portrait
            out += pclEscStr("&l0O")
            // Raster resolution
            out += pclEscStr("*t\(resolution)R")
            // Source width
            out += pclEscStr("*r\(page.widthPx)S")
            // Start raster
            out += pclEscStr("*r0A")

            for row in 0..<page.heightPx {
                let start = row * rowBytes
                let scanline = Array(rgb[start..<(start + rowBytes)])
                // Compression mode 0
                out += pclEscStr("*b0M")
                // Transfer row
                out += pclEscStr("*b\(rowBytes)W")
                out += scanline
            }

            // End raster
            out += pclEscStr("*rC")
            // Eject page
            out += pclEscStr("&l0H")
        }

        out += [0x1B, 0x45]  // ESC E reset
        return Data(out)
    }

    // MARK: - PCL XL

    /// Exports `pdf` as PCL XL (binary, little-endian binding, Protocol 3.0).
    @DocRenderActor
    public static func exportPCLXL(pdf: Data, resolution: Int = 300) throws -> Data {
        let pages = try renderPages(pdf: pdf, resolution: resolution)
        var out = [UInt8]()

        // Stream header (ASCII, CRLF-terminated)
        let hdr1 = "\u{1B}%-12345X@PJL\r\n"
        let hdr2 = ") HP-PCL XL;3;0;Comment DocReader\r\n"
        out += Array(hdr1.utf8)
        out += Array(hdr2.utf8)

        // BeginSession (0x41)
        out += pclxlUByte(1)          ; out += pclxlAttr(0x35)  // MeasureType = Inch
        out += pclxlUInt16(UInt16(resolution)) ; out += pclxlAttr(0x36)  // UnitsPerMeasure
        out += pclxlUByte(2)          ; out += pclxlAttr(0x34)  // ErrorReport = BackChannel
        out += [0x41]

        for page in pages {
            let mediaSz = UInt8(page.widthPt > 600 ? 4 : 3)  // 4=Letter, 3=A4
            let wPx = UInt16(min(page.widthPx,  Int(UInt16.max)))
            let hPx = UInt16(min(page.heightPx, Int(UInt16.max)))

            // BeginPage (0x43)
            out += pclxlUByte(mediaSz) ; out += pclxlAttr(0x25)  // MediaSize
            out += pclxlUByte(0)       ; out += pclxlAttr(0x28)  // Orientation = portrait
            out += [0x43]

            // BeginImage (0xB0)
            out += pclxlUByte(2)   ; out += pclxlAttr(0x03)  // ColorSpace = eRGB
            out += pclxlUByte(8)   ; out += pclxlAttr(0x02)  // PaletteDepth = 8
            out += pclxlUInt16(wPx) ; out += pclxlAttr(0x23)  // SourceWidth
            out += pclxlUInt16(hPx) ; out += pclxlAttr(0x24)  // SourceHeight
            out += pclxlUInt16XY(wPx, hPx) ; out += pclxlAttr(0x0D)  // DestinationSize
            out += [0xB0]

            // ReadImage (0xB1)
            let totalBytes = UInt32(page.widthPx) * UInt32(page.heightPx) * 3
            out += pclxlUInt16(0)           ; out += pclxlAttr(0x49)  // StartLine
            out += pclxlUInt16(hPx)         ; out += pclxlAttr(0x4A)  // BlockHeight
            out += pclxlUByte(0)            ; out += pclxlAttr(0x05)  // CompressMode = none
            out += pclxlUByte(0)            ; out += pclxlAttr(0x06)  // DataOrg = binary
            out += pclxlUInt32(totalBytes)  ; out += pclxlAttr(0x44)  // DataLength
            out += [0xB1]
            out += [UInt8](page.rgbData)

            // EndImage (0xB2)
            out += [0xB2]

            // EndPage (0x44)
            out += pclxlUInt16(1) ; out += pclxlAttr(0x31)  // PageCopies
            out += [0x44]
        }

        // EndSession (0x42)
        out += [0x42]
        return Data(out)
    }

    // MARK: - Private helpers

    private static func writeBEUInt32(_ v: UInt32, into buf: inout [UInt8], at offset: Int) {
        buf[offset]     = UInt8((v >> 24) & 0xFF)
        buf[offset + 1] = UInt8((v >> 16) & 0xFF)
        buf[offset + 2] = UInt8((v >>  8) & 0xFF)
        buf[offset + 3] = UInt8( v        & 0xFF)
    }

    private static func writeBEUInt32Into(_ buf: inout [UInt8], _ v: UInt32) {
        buf.append(UInt8((v >> 24) & 0xFF))
        buf.append(UInt8((v >> 16) & 0xFF))
        buf.append(UInt8((v >>  8) & 0xFF))
        buf.append(UInt8( v        & 0xFF))
    }

    private static func writeFloat32BE(_ v: Float, into buf: inout [UInt8], at offset: Int) {
        let bits = v.bitPattern
        buf[offset]     = UInt8((bits >> 24) & 0xFF)
        buf[offset + 1] = UInt8((bits >> 16) & 0xFF)
        buf[offset + 2] = UInt8((bits >>  8) & 0xFF)
        buf[offset + 3] = UInt8( bits        & 0xFF)
    }

    /// PackBits row encoder (pixel-granularity, 3 bytes/pixel).
    static func packBitsRow(_ row: ArraySlice<UInt8>) -> [UInt8] {
        let pixels = Array(row)
        let pixelCount = pixels.count / 3
        var out = [UInt8]()
        var i = 0
        while i < pixelCount {
            let r = pixels[i * 3], g = pixels[i * 3 + 1], b = pixels[i * 3 + 2]
            // Count run of identical pixels
            var runLen = 1
            while runLen < 128 && (i + runLen) < pixelCount {
                let j = i + runLen
                if pixels[j * 3] == r && pixels[j * 3 + 1] == g && pixels[j * 3 + 2] == b {
                    runLen += 1
                } else { break }
            }
            if runLen > 1 {
                out.append(0x80 | UInt8(128 - runLen))
                out += [r, g, b]
                i += runLen
            } else {
                // Count literal run (pixels that differ from their successor)
                var litLen = 1
                while litLen < 128 && (i + litLen) < pixelCount {
                    let j = i + litLen
                    let nr = pixels[j * 3], ng = pixels[j * 3 + 1], nb = pixels[j * 3 + 2]
                    if j + 1 < pixelCount {
                        let nr2 = pixels[(j + 1) * 3], ng2 = pixels[(j + 1) * 3 + 1], nb2 = pixels[(j + 1) * 3 + 2]
                        if nr == nr2 && ng == ng2 && nb == nb2 { break }  // next pair is a run
                    }
                    _ = (nr, ng, nb)
                    litLen += 1
                }
                out.append(UInt8(litLen - 1))
                for k in 0..<litLen {
                    let j = i + k
                    out += [pixels[j * 3], pixels[j * 3 + 1], pixels[j * 3 + 2]]
                }
                i += litLen
            }
        }
        return out
    }

    private static func pclPaperCode(widthPt: CGFloat) -> Int {
        let w = Int(widthPt.rounded())
        if w == 612 { return 2 }   // Letter
        if w == 595 { return 26 }  // A4
        return 101                  // custom
    }

    private static func pclEscStr(_ s: String) -> [UInt8] {
        [0x1B] + Array(s.utf8)
    }

    private static func pclxlUByte(_ v: UInt8) -> [UInt8]  { [0xC0, v] }
    private static func pclxlUInt16(_ v: UInt16) -> [UInt8] {
        [0xC1, UInt8(v & 0xFF), UInt8(v >> 8)]
    }
    private static func pclxlUInt32(_ v: UInt32) -> [UInt8] {
        [0xC2,
         UInt8( v        & 0xFF),
         UInt8((v >>  8) & 0xFF),
         UInt8((v >> 16) & 0xFF),
         UInt8((v >> 24) & 0xFF)]
    }
    private static func pclxlUInt16XY(_ x: UInt16, _ y: UInt16) -> [UInt8] {
        [0xD1, UInt8(x & 0xFF), UInt8(x >> 8), UInt8(y & 0xFF), UInt8(y >> 8)]
    }
    private static func pclxlAttr(_ id: UInt8) -> [UInt8] { [0xF8, id] }
}
