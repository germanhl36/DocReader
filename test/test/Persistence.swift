//
//  Persistence.swift
//  test
//
//  Created by German Huerta on 28/02/26.
//

import CoreData
import DocReader

struct PersistenceController {
    static let shared = PersistenceController()

    @MainActor
    static let preview: PersistenceController = {
        let result = PersistenceController(inMemory: true)
        let viewContext = result.container.viewContext
        for _ in 0..<10 {
            let newItem = Item(context: viewContext)
            newItem.timestamp = Date()
        }
        do {
            try viewContext.save()
        } catch {
            // Replace this implementation with code to handle the error appropriately.
            // fatalError() causes the application to generate a crash log and terminate. You should not use this function in a shipping application, although it may be useful during development.
            let nsError = error as NSError
            fatalError("Unresolved error \(nsError), \(nsError.userInfo)")
        }
        return result
    }()

    let container: NSPersistentContainer

    init(inMemory: Bool = false) {
        container = NSPersistentContainer(name: "test")
        if inMemory {
            container.persistentStoreDescriptions.first!.url = URL(fileURLWithPath: "/dev/null")
        }
        container.loadPersistentStores(completionHandler: { (storeDescription, error) in
            if let error = error as NSError? {
                // Replace this implementation with code to handle the error appropriately.
                // fatalError() causes the application to generate a crash log and terminate. You should not use this function in a shipping application, although it may be useful during development.

                /*
                 Typical reasons for an error here include:
                 * The parent directory does not exist, cannot be created, or disallows writing.
                 * The persistent store is not accessible, due to permissions or data protection when the device is locked.
                 * The device is out of space.
                 * The store could not be migrated to the current model version.
                 Check the error message to determine what the actual problem was.
                 */
                fatalError("Unresolved error \(error), \(error.userInfo)")
            }
        })
        container.viewContext.automaticallyMergesChangesFromParent = true
    }
}
// MARK: - Document Reader

struct DocumentReader {
    let url: URL
    
    init(url: URL) {
        self.url = url
    }
    
    /// Check if the URL points to a supported document file
    var isSupported: Bool {
        DocReader.isSupported(url: url)
    }
    
    /// Open and process a document, returning metadata and page information
    func processDocument() async throws -> DocumentInfo? {
        // Check if a URL points to a supported file
        guard DocReader.isSupported(url: url) else { return nil }
        
        // Open the document
        let doc = try await DocReader.open(url: url)
        
        // Inspect
        let pages = try await doc.pageCount
        let size = try await doc.pageSize(at: 0)
        let metadata = try await doc.metadata
        
        return DocumentInfo(
            pageCount: pages,
            firstPageSize: size,
            metadata: metadata,
            document: doc
        )
    }
    
    /// Export the entire document to PDF
    func exportToPDF() async throws -> Data? {
        guard DocReader.isSupported(url: url) else { return nil }
        
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF()
        
        return pdf
    }
    
    /// Export a specific page range to PDF
    func exportPagesToPDF(range: ClosedRange<Int>) async throws -> Data? {
        guard DocReader.isSupported(url: url) else { return nil }
        
        let doc = try await DocReader.open(url: url)
        let pdf = try await doc.exportPDF(pages: range)
        
        return pdf
    }
    
    /// Export a single page to PDF
    func exportSinglePageToPDF(pageIndex: Int) async throws -> Data? {
        guard DocReader.isSupported(url: url) else { return nil }
        
        let doc = try await DocReader.open(url: url)
        let page = try await doc.exportPDF(pages: pageIndex...pageIndex)
        
        return page
    }
}

// MARK: - Document Info

struct DocumentInfo {
    let pageCount: Int
    let firstPageSize: CGSize
    let metadata: DocumentMetadata
    let document: Any // The DocReader document instance
}

