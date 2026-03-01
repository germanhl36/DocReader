# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Build
swift build
swift build -c release

# Test (all)
swift test --parallel

# Test (single test)
swift test --filter DocReaderTests/testOpenDocx
swift test --filter ContentModelTests/testColumnLabelConversion

# Generate test fixtures (requires Python 3 + openpyxl/python-pptx)
python3 scripts/generate_fixtures.py

# Generate documentation
swift package generate-documentation
```

## Architecture

### Entry point → Parser → Renderer pipeline

```
DocReader.open(url:)
  └─ DocReaderFactory          picks parser by DocumentFormat
       ├─ OOXMLWordParser       .docx — ZIP + XMLParser SAX
       ├─ OOXMLSheetParser      .xlsx — ZIP + XMLParser SAX
       ├─ OOXMLSlideParser      .pptx — ZIP + XMLParser SAX
       ├─ DOCLegacyParser       .doc  — OLEKit streams
       ├─ XLSLegacyParser       .xls  — OLEKit streams
       └─ PPTLegacyParser       .ppt  — OLEKit streams

PDFExporter.export(parser:pages:)
  └─ casts parser to DocContentProviding (internal protocol)
       ├─ buildPageContents() → [PageContent]
       └─ routes to WordPageRenderer / SheetPageRenderer / SlidePageRenderer
            └─ CoreText CTFramesetter + CGContext (no UIKit)
```

### Key protocols

- **`DocReadable`** (public, actor protocol) — the only type callers see. Properties: `url`, `format`, `pageCount`, `pageSize(at:)`, `metadata`, `exportPDF()`.
- **`DocContentProviding`** (internal, actor protocol) — implemented by all three OOXML parsers. `buildPageContents() throws -> [PageContent]`. Legacy parsers use placeholder rendering and do NOT implement this.
- **`DocRenderActor`** (`@globalActor`) — all CoreGraphics drawing must be isolated to this actor.

### Content model (`RenderModels.swift`)

All `Sendable` value types passed from parsers to renderers:
- Word: `WordRunContent` → `WordParagraphContent` → `WordPageContent` (with `WordPageMargins`)
- Sheet: `SheetCellContent` → `SheetPageContent`
- Slide: `SlideTextBoxContent` → `SlidePageContent`
- Union: `enum PageContent { case word, sheet, slide }`

Page breaks in Word documents are emitted as sentinel paragraphs with `styleName == "__pagebreak__"`. `splitIntoWordPages()` splits the flat array at these sentinels.

### OOXML parsing

All three OOXML parsers use SAX (`XMLParserDelegate`) state machines — no DOM, no third-party XML library. Key paths parsed:
- Word: `word/document.xml`, `docProps/app.xml`, `docProps/core.xml`
- Sheet: `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/sharedStrings.xml`, `xl/worksheets/sheetN.xml`
- Slide: `ppt/presentation.xml`, `ppt/_rels/presentation.xml.rels`, `ppt/slides/slideN.xml`

`OOXMLZipExtractor` wraps ZIPFoundation; call `validateOOXMLStructure()` before extracting entries.

### Legacy (OLE2) parsing

Uses OLEKit. API: `OLEFile(_ path: String)` (not `url`), navigate via `ole.root.children`, read with `ole.stream(_ entry:)`. All streams must be ≥ 4096 bytes to avoid the mini-stream path.

`OLEPropertySetReader` decodes MS-OLEPS property set streams (title, author, dates). FILETIME values are split into `lo`/`hi` halves before combining to avoid Swift type-checker timeouts in release builds.

## Swift 6 concurrency rules

`StrictConcurrency = complete` is enforced. Every parser is an `actor`. Shared state across async boundaries must be `Sendable`. All renderers are `enum` with `@DocRenderActor` methods. When adding new parsers or renderers, follow the same pattern.

## Jira

Project key: `DOCR`. Epics use type `Tarea`; stories use type `Subtarea` with `parent: {key: "DOCR-xx"}` in the POST body. Current epics: DOCR-1..5 (v1.0), DOCR-39 (Iteration 6).
