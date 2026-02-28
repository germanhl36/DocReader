# ``DocReader``

Read and export Microsoft Office documents natively on iOS.

## Overview

DocReader is a Swift Package that lets you open `.doc`, `.docx`, `.xls`, `.xlsx`,
`.ppt`, and `.pptx` files directly on iOS — without any server round-trips or
third-party cloud services.

### Key features

- **Format detection** — automatic detection by file extension and magic bytes.
- **Page count & dimensions** — read page count and size in points.
- **Document metadata** — title, author, creation and modification dates.
- **PDF export** — export full documents or page subsets to `Data` using CoreGraphics.
- **Swift 6 concurrency** — all operations are `async`/`await`, safely isolated with `@DocRenderActor`.

## Getting started

```swift
import DocReader

// 1. Check if a file is supported
guard DocReader.isSupported(url: fileURL) else { return }

// 2. Open the document
let document = try await DocReader.open(url: fileURL)

// 3. Read page count
let pages = try await document.pageCount

// 4. Export to PDF
let pdfData = try await document.exportPDF()
```

## Topics

### Entry Point

- ``DocReader``

### Protocol

- ``DocReadable``

### Document Information

- ``DocumentFormat``
- ``DocumentMetadata``

### Error Handling

- ``DocReaderError``

### Rendering

- ``DocRenderActor``

### Articles

- <doc:ReadingYourFirstDocument>
