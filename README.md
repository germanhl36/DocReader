# DocReader

A native Swift Package for reading and inspecting Microsoft Office documents on iOS.

[![CI](https://github.com/YOUR_ORG/DocReader/actions/workflows/ci.yml/badge.svg)](https://github.com/YOUR_ORG/DocReader/actions/workflows/ci.yml)
[![Swift 6](https://img.shields.io/badge/Swift-6.0-orange)](https://swift.org)
[![iOS 16+](https://img.shields.io/badge/iOS-16%2B-blue)](https://developer.apple.com/ios/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green)](LICENSE)

## Features

| Capability | Formats |
|---|---|
| Page count | `.docx` `.xlsx` `.pptx` `.doc` `.xls` `.ppt` |
| Page dimensions | `.docx` `.xlsx` `.pptx` `.ppt` |
| Document metadata | All formats |
| PDF export | All formats |

- Swift 6 strict concurrency (`StrictConcurrency = complete`)
- Off-main-thread rendering via `@DocRenderActor`
- No network access at runtime
- Zero UIKit dependency (CoreGraphics only)
- Localized errors in English and Spanish

## Installation

### Swift Package Manager

```swift
// Package.swift
dependencies: [
    .package(url: "https://github.com/YOUR_ORG/DocReader.git", from: "1.0.0")
]
```

### CocoaPods

```ruby
pod 'DocReader', '~> 1.0'
```

## Quick start

```swift
import DocReader

// Check if a URL points to a supported file
guard DocReader.isSupported(url: url) else { return }

// Open the document
let doc = try await DocReader.open(url: url)

// Inspect
let pages    = try await doc.pageCount
let size     = try await doc.pageSize(at: 0)
let metadata = try await doc.metadata

// Export to PDF
let pdf = try await doc.exportPDF()

// Export a single page
let page1 = try await doc.exportPDF(pages: 0...0)
```

## Architecture

```
Public API      → DocReadable protocol + public Swift types (DocC documented)
Document Engine → DocReader (entry point), DocReaderFactory, DocFormatDetector
Format Parsers  → OOXML (ZIPFoundation + XMLParser) | Legacy (OLEKit)
Rendering Layer → CGPDFContextCreate + CoreText (CGContext — no UIKit)
Storage Adapter → FileManager / sandbox (Foundation only)
```

## Requirements

- iOS 16+ / macOS 13+
- Swift 6.0 / Xcode 16+

## Generating fixtures

Integration tests expect fixture files in `Tests/DocReaderIntegrationTests/Fixtures/`.
In CI, fixtures are generated via LibreOffice headless. To generate them locally:

```bash
libreoffice --headless --convert-to docx tests/fixtures/source/*.odt --outdir Tests/DocReaderIntegrationTests/Fixtures/
```

Fixture files must be ≤ 500 KB each.

## License

MIT — see [LICENSE](LICENSE).
