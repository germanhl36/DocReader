# Changelog

All notable changes to DocReader are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] — 2026-02-27

### Added

- **DocReader.open(url:)** — async entry point returning a `DocReadable` parser
- **DocReader.isSupported(url:)** — synchronous extension check for 6 formats
- **DocReadable** protocol — `pageCount`, `pageSize(at:)`, `metadata`, `exportPDF()`, `exportPDF(pages:)`
- **DocumentFormat** — enum covering `.docx`, `.xlsx`, `.pptx`, `.doc`, `.xls`, `.ppt` with `family` and `isLegacy`
- **DocumentMetadata** — `Sendable` struct: title, author, created/modified dates
- **DocReaderError** — 6 typed errors with `LocalizedError` in English and Spanish
- **DocFormatDetector** — extension-based + magic-byte detection (ZIP and OLE2)
- **OOXMLWordParser** — parses `.docx` (page count from `docProps/app.xml <Pages>`, page size from `<w:pgSz>`)
- **OOXMLSheetParser** — parses `.xlsx` (sheet count from `xl/workbook.xml <sheet>`)
- **OOXMLSlideParser** — parses `.pptx` (slide count from `ppt/presentation.xml <p:sldId>`, EMU → points)
- **DOCLegacyParser** — reads `.doc` via OLEKit + OLE property set reader
- **XLSLegacyParser** — counts sheets by scanning BIFF8 BOUNDSHEET records
- **PPTLegacyParser** — counts slides by scanning SlideContainer atoms (0x03E8)
- **OLEPropertySetReader** — minimal OLE2 property set decoder ([MS-OLEPS])
- **DocRenderActor** — `@globalActor` for off-main-thread PDF generation
- **PDFExporter** — `CGPDFContextCreate`-based PDF output with cancellation support
- **WordPageRenderer**, **SheetPageRenderer**, **SlidePageRenderer** — CoreText-based page renderers
- DocC documentation with tutorial "Reading Your First Document"
- GitHub Actions CI (build + test + coverage ≥ 80%) and release (XCFramework) workflows
- CocoaPods `.podspec`
- Localization: `Localizable.xcstrings` (en + es)

### Dependencies

- [ZIPFoundation](https://github.com/weichsel/ZIPFoundation) 0.9+
- [OLEKit](https://github.com/CoreOffice/OLEKit) 0.2.x

---

[1.0.0]: https://github.com/YOUR_ORG/DocReader/releases/tag/v1.0.0
