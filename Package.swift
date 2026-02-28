// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "DocReader",
    platforms: [
        .iOS(.v16),
        .macOS(.v13),
    ],
    products: [
        .library(
            name: "DocReader",
            targets: ["DocReader"]
        ),
    ],
    dependencies: [
        .package(
            url: "https://github.com/weichsel/ZIPFoundation.git",
            .upToNextMajor(from: "0.9.0")
        ),
        .package(
            url: "https://github.com/CoreOffice/OLEKit.git",
            .upToNextMinor(from: "0.2.0")
        ),
        .package(
            url: "https://github.com/swiftlang/swift-docc-plugin",
            from: "1.0.0"
        ),
    ],
    targets: [
        .target(
            name: "DocReader",
            dependencies: [
                .product(name: "ZIPFoundation", package: "ZIPFoundation"),
                .product(name: "OLEKit", package: "OLEKit"),
            ],
            resources: [
                .process("Resources"),
            ],
            swiftSettings: [
                .enableExperimentalFeature("StrictConcurrency"),
            ]
        ),
        .testTarget(
            name: "DocReaderTests",
            dependencies: ["DocReader"],
            swiftSettings: [
                .enableExperimentalFeature("StrictConcurrency"),
            ]
        ),
        .testTarget(
            name: "DocReaderIntegrationTests",
            dependencies: ["DocReader"],
            resources: [
                .copy("Fixtures"),
            ],
            swiftSettings: [
                .enableExperimentalFeature("StrictConcurrency"),
            ]
        ),
    ]
)
