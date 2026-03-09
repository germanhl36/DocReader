//
//  ContentView.swift
//  test
//
//  Created by German Huerta on 28/02/26.
//

import SwiftUI
import CoreData
import PDFKit
import DocReader

struct ContentView: View {
    @Environment(\.managedObjectContext) private var viewContext
    @State private var pdfData: Data?
    @State private var isLoading = false
    @State private var errorMessage: String?
    @State private var documentInfo: DocumentInfo?

    var body: some View {
        NavigationView {
            VStack(spacing: 20) {
                // Header with document info
                if let docInfo = documentInfo {
                    DocumentInfoCard(documentInfo: docInfo)
                        .padding()
                }
                
                // Main content area
                if isLoading {
                    VStack(spacing: 16) {
                        ProgressView()
                            .scaleEffect(1.5)
                        Text("Processing document...")
                            .font(.headline)
                    }
                } else if let pdfData = pdfData {
                    PDFViewerWrapper(data: pdfData)
                } else {
                    VStack(spacing: 16) {
                        Image(systemName: "doc.text.magnifyingglass")
                            .font(.system(size: 60))
                            .foregroundStyle(.secondary)
                        
                        Text("No PDF loaded")
                            .font(.title2)
                            .foregroundStyle(.secondary)
                        
                        Text("Press the button below to load DocReader_SDD.docx")
                            .font(.caption)
                            .foregroundStyle(.tertiary)
                            .multilineTextAlignment(.center)
                            .padding(.horizontal)
                    }
                }
                
                // Error message if any
                if let error = errorMessage {
                    Text(error)
                        .foregroundStyle(.red)
                        .padding()
                        .background(Color.red.opacity(0.1))
                        .cornerRadius(8)
                        .padding(.horizontal)
                }
                
                Spacer()
                
                // Load button
                Button(action: loadDocument) {
                    HStack {
                        Image(systemName: "arrow.down.doc")
                        Text("Load DocReader_SDD.docx")
                    }
                    .font(.headline)
                    .foregroundStyle(.white)
                    .frame(maxWidth: .infinity)
                    .padding()
                    .background(Color.accentColor)
                    .cornerRadius(12)
                }
                .padding()
                .disabled(isLoading)
            }
            .navigationTitle("Document Viewer")
            .task { loadDocument() }
            .toolbar {
                ToolbarItem(placement: .navigationBarTrailing) {
                    if pdfData != nil {
                        Button(action: clearDocument) {
                            Label("Clear", systemImage: "trash")
                        }
                    }
                }
            }
        }
    }
    
    private func loadDocument() {
        isLoading = true
        errorMessage = nil
        documentInfo = nil
        
        Task {
            do {
                // Get the document URL from the bundle
                guard let docURL = Bundle.main.url(forResource: "DocReader_SDD", withExtension: "pdf") else {
                    await MainActor.run {
                        errorMessage = "DocReader_SDD.pdf not found in bundle"
                        isLoading = false
                    }
                    return
                }
                
                // Create DocumentReader instance
//                let reader = DocumentReader(url: docURL)
                
                // Check if the document is supported
//                guard reader.isSupported else {
//                    await MainActor.run {
//                        errorMessage = "Document format not supported"
//                        isLoading = false
//                    }
//                    return
//                }
                
                // Process document to get info
//                let docInfo = try await reader.processDocument()
                
                // Export to PDF
//                guard let pdf = try await reader.exportToPDF() else {
//                    await MainActor.run {
//                        errorMessage = "Failed to export PDF"
//                        isLoading = false
//                    }
//                    return
//                }
                
                // Save PDF to documents folder
                //let savedURL = try saveToDocuments(pdfData: pdf, filename: "DocReader_SDD.pdf")
                
                // Print the URL to console
                //print("✅ PDF saved successfully to: \(savedURL.path)")
                //print("📄 File URL: \(savedURL.absoluteString)")
                
                guard let pdf = try? Data(contentsOf: docURL) else {
                    await MainActor.run {
                        errorMessage = "Failed to export PDF"
                        isLoading = false
                    }
                    return
                }
                // PWG (CUPS / IPP printer
                
                let pwgFromPDF = try await PrintExporter.exportPWGRaster(pdf: pdf, resolution: 300)
                let savedPWGURL = try saveToDocuments(pdfData: pwgFromPDF, filename: "DocReader_SDD.pwg")

                  // URF / UNIRAST (AirPrint)
                let urfFromPDF  = try await PrintExporter.exportURF(pdf: pdf, resolution: 300)
                let savedURFURL = try saveToDocuments(pdfData: urfFromPDF, filename: "DocReader_SDD.urf")


                  // PCL 5 (LaserJet-compatible printers)
//                let pcl  = try await PrintExporter.exportPCL(pdf: pdf, resolution: 300)
//
//                let savedPCLURL = try saveToDocuments(pdfData: pcl, filename: "DocReader_SDD.pcl")
//
//                  // PCL XL / PCL 6 (newer HP and compatible printers)
//                  let pclxl = try await PrintExporter.exportPCLXL(pdf: pdf, resolution: 600)
//                let savedPCLXLURL = try saveToDocuments(pdfData: pclxl, filename: "DocReader_SDD_PCLXL.pcl")

                await MainActor.run {
//                    self.documentInfo = docInfo
                    self.pdfData = pdf
                    self.isLoading = false
                }
            } catch {
                await MainActor.run {
                    errorMessage = "Error: \(error.localizedDescription)"
                    isLoading = false
                }
            }
        }
    }
    
    private func saveToDocuments(pdfData: Data, filename: String) throws -> URL {
        // Get the documents directory
        let documentsDirectory = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask)[0]
        
        // Create the file URL
        let fileURL = documentsDirectory.appendingPathComponent(filename)
        
        // Write the PDF data to the file
        try pdfData.write(to: fileURL, options: .atomic)
        
        return fileURL
    }
    
    
    
    private func clearDocument() {
        withAnimation {
            pdfData = nil
            documentInfo = nil
            errorMessage = nil
        }
    }
}

// MARK: - Document Info Card

struct DocumentInfoCard: View {
    let documentInfo: DocumentInfo
    
    var body: some View {
        VStack(alignment: .leading, spacing: 8) {
            Text("Document Information")
                .font(.headline)
            
            HStack {
                Label("\(documentInfo.pageCount)", systemImage: "doc.text")
                Text("pages")
                    .foregroundStyle(.secondary)
            }
            
            HStack {
                Label("Size:", systemImage: "ruler")
                Text("\(Int(documentInfo.firstPageSize.width)) × \(Int(documentInfo.firstPageSize.height)) pt")
                    .foregroundStyle(.secondary)
            }
        }
        .frame(maxWidth: .infinity, alignment: .leading)
        .padding()
        .background(Color(.systemGray6))
        .cornerRadius(12)
    }
}

// MARK: - PDF Viewer Wrapper

struct PDFViewerWrapper: View {
    let data: Data
    
    var body: some View {
        PDFKitView(data: data)
    }
}

// MARK: - PDFKit View

struct PDFKitView: UIViewRepresentable {
    let data: Data
    
    func makeUIView(context: Context) -> PDFView {
        let pdfView = PDFView()
        pdfView.autoScales = true
        pdfView.displayMode = .singlePageContinuous
        pdfView.displayDirection = .vertical
        return pdfView
    }
    
    func updateUIView(_ pdfView: PDFView, context: Context) {
        if let document = PDFDocument(data: data) {
            pdfView.document = document
        }
    }
}

private let itemFormatter: DateFormatter = {
    let formatter = DateFormatter()
    formatter.dateStyle = .short
    formatter.timeStyle = .medium
    return formatter
}()

#Preview {
    ContentView().environment(\.managedObjectContext, PersistenceController.preview.container.viewContext)
}
