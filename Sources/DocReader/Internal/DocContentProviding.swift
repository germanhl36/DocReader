import Foundation

/// Internal protocol for OOXML parsers that can provide structured page content for rendering.
/// Not exposed in the public API.
protocol DocContentProviding: Actor {
    func buildPageContents() async throws -> [PageContent]
}
