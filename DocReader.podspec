Pod::Spec.new do |s|
  s.name             = 'DocReader'
  s.version          = '1.0.0'
  s.summary          = 'Read and export Microsoft Office documents (.doc, .docx, .xls, .xlsx, .ppt, .pptx) to PDF on iOS.'
  s.description      = <<-DESC
    DocReader is a native Swift library for reading and inspecting Microsoft Office
    documents. It supports OOXML formats (docx, xlsx, pptx) and legacy binary formats
    (doc, xls, ppt). Documents can be inspected for page count, page dimensions,
    and metadata, and exported to PDF using CoreGraphics — no UIKit required.
  DESC

  s.homepage         = 'https://github.com/YOUR_ORG/DocReader'
  s.license          = { :type => 'MIT', :file => 'LICENSE' }
  s.author           = { 'Your Name' => 'your@email.com' }
  s.source           = { :git => 'https://github.com/YOUR_ORG/DocReader.git', :tag => s.version.to_s }

  s.ios.deployment_target = '16.0'
  s.swift_version = '6.0'

  s.source_files = 'Sources/DocReader/**/*.swift'
  s.resource_bundles = {
    'DocReader' => ['Sources/DocReader/Resources/**/*']
  }

  s.frameworks = 'Foundation', 'CoreGraphics', 'CoreText', 'PDFKit'

  # ZIPFoundation dependency
  s.dependency 'ZIPFoundation', '~> 0.9'

  # OLEKit: if no official podspec is available, vendor the sources
  # Vendored sources are placed under Sources/DocReader/Vendor/OLEKit/
  # s.dependency 'OLEKit', '~> 0.2'
  # Uncomment below if vendoring:
  # s.source_files = 'Sources/DocReader/**/*.swift', 'Sources/DocReader/Vendor/OLEKit/**/*.swift'
end
