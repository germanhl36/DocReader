# Reading Your First Document

Learn how to open a Microsoft Office document and export it to PDF in just a few lines of Swift.

## Overview

This article walks through the complete flow: checking support, opening a document,
reading its metadata, and exporting it to a PDF file.

## Step 1: Add DocReader to your project

Add the package in Xcode via **File → Add Package Dependencies**:

```
https://github.com/YOUR_ORG/DocReader
```

Or add it to your `Package.swift`:

```swift
dependencies: [
    .package(url: "https://github.com/YOUR_ORG/DocReader.git", from: "1.0.0")
]
```

## Step 2: Check if a file is supported

```swift
import DocReader

let fileURL = // ... URL to an Office document
guard DocReader.isSupported(url: fileURL) else {
    print("File format not supported")
    return
}
```

``DocReader/isSupported(url:)`` checks the file extension synchronously without
touching the disk.

## Step 3: Open the document

```swift
do {
    let document = try await DocReader.open(url: fileURL)
    // document conforms to DocReadable
} catch DocReaderError.fileNotFound {
    print("File does not exist")
} catch DocReaderError.unsupportedFormat {
    print("Unsupported format")
} catch DocReaderError.corruptedFile {
    print("File is corrupted")
}
```

## Step 4: Read page count and metadata

```swift
let pageCount = try await document.pageCount
let pageSize  = try await document.pageSize(at: 0)
let meta      = try await document.metadata

print("Title: \(meta.title ?? "—")")
print("Pages: \(pageCount), size: \(pageSize.width)×\(pageSize.height) pt")
```

## Step 5: Export to PDF

```swift
// Export all pages
let fullPDF = try await document.exportPDF()

// Export only the first page
let firstPage = try await document.exportPDF(pages: 0...0)

// Save to disk
let outputURL = FileManager.default.temporaryDirectory
    .appendingPathComponent("output.pdf")
try firstPage.write(to: outputURL)
```

## Handling cancellation

Export operations respect Swift structured concurrency cancellation:

```swift
let task = Task {
    let pdf = try await document.exportPDF()
    return pdf
}

// Cancel mid-export
task.cancel()

do {
    let _ = try await task.value
} catch DocReaderError.exportCancelled {
    print("Export was cancelled cleanly")
}
```
