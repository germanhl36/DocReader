import Foundation

/// Minimal reader for OLE2 property set streams (e.g. SummaryInformation).
///
/// The OLE property set format is defined in [MS-OLEPS].
/// This implementation handles the subset needed for page count, title, and author.
enum OLEPropertySetReader {
    // MARK: - Property types (VARTYPE)
    private static let VT_I2: UInt16 = 2
    private static let VT_I4: UInt16 = 3
    private static let VT_LPSTR: UInt16 = 30
    private static let VT_FILETIME: UInt16 = 64

    // MARK: - Public API

    /// Reads an integer property by `propertyID` from a property set stream blob.
    static func readInt(from data: Data, propertyID: UInt32) -> Int? {
        guard let offset = findPropertyOffset(in: data, propertyID: propertyID) else { return nil }
        return readIntValue(from: data, at: offset)
    }

    /// Reads a string property by `propertyID` from a property set stream blob.
    static func readString(from data: Data, propertyID: UInt32) -> String? {
        guard let offset = findPropertyOffset(in: data, propertyID: propertyID) else { return nil }
        return readStringValue(from: data, at: offset)
    }

    /// Reads a FILETIME property by `propertyID` from a property set stream blob.
    ///
    /// FILETIME is a 64-bit value representing the number of 100-nanosecond intervals
    /// since January 1, 1601 (Windows epoch). Returns `nil` if the property is absent
    /// or the stored value is zero (unset).
    static func readDate(from data: Data, propertyID: UInt32) -> Date? {
        guard let offset = findPropertyOffset(in: data, propertyID: propertyID) else { return nil }
        return readFileTimeValue(from: data, at: offset)
    }

    // MARK: - Implementation

    /// Locates the byte offset of a property value in the property set stream.
    ///
    /// Property Set Stream layout ([MS-OLEPS] §2.2):
    /// - Header: 28 bytes (byte order, version, FMTID...)
    /// - PropertySetListHeader (4 bytes: count, then reservedOffset = 0)
    ///   Actually: NumProperties (4 bytes) at offset 0x28
    ///   Then: PropertyIdentifierAndOffset pairs (8 bytes each): propID (4) + offset (4)
    ///   Offsets are relative to start of PropertySet section.
    ///
    /// We detect the properties section start by scanning the known header.
    private static func findPropertyOffset(in data: Data, propertyID: UInt32) -> Int? {
        let bytes = [UInt8](data)
        guard bytes.count >= 52 else { return nil }

        // PropertySet section offset is at bytes 44..47 in the stream header
        let sectionOffset = Int(
            UInt32(bytes[44])
            | (UInt32(bytes[45]) << 8)
            | (UInt32(bytes[46]) << 16)
            | (UInt32(bytes[47]) << 24)
        )

        guard sectionOffset + 8 <= bytes.count else { return nil }

        // PropertySet starts with: Size (4) + NumProperties (4)
        let numProperties = Int(
            UInt32(bytes[sectionOffset + 4])
            | (UInt32(bytes[sectionOffset + 5]) << 8)
            | (UInt32(bytes[sectionOffset + 6]) << 16)
            | (UInt32(bytes[sectionOffset + 7]) << 24)
        )

        guard numProperties > 0, numProperties < 10_000 else { return nil }

        // Property identifier + offset array starts at sectionOffset + 8
        let idOffsetBase = sectionOffset + 8

        for i in 0..<numProperties {
            let entryBase = idOffsetBase + i * 8
            guard entryBase + 8 <= bytes.count else { return nil }

            let pid = UInt32(bytes[entryBase])
                    | (UInt32(bytes[entryBase + 1]) << 8)
                    | (UInt32(bytes[entryBase + 2]) << 16)
                    | (UInt32(bytes[entryBase + 3]) << 24)

            if pid == propertyID {
                let relativeOffset = Int(
                    UInt32(bytes[entryBase + 4])
                    | (UInt32(bytes[entryBase + 5]) << 8)
                    | (UInt32(bytes[entryBase + 6]) << 16)
                    | (UInt32(bytes[entryBase + 7]) << 24)
                )
                return sectionOffset + relativeOffset
            }
        }
        return nil
    }

    private static func readIntValue(from data: Data, at offset: Int) -> Int? {
        let bytes = [UInt8](data)
        guard offset + 6 <= bytes.count else { return nil }
        let vt = UInt16(bytes[offset]) | (UInt16(bytes[offset + 1]) << 8)
        switch vt {
        case VT_I2:
            guard offset + 4 <= bytes.count else { return nil }
            return Int(Int16(bitPattern: UInt16(bytes[offset + 2]) | (UInt16(bytes[offset + 3]) << 8)))
        case VT_I4:
            guard offset + 6 <= bytes.count else { return nil }
            let raw = UInt32(bytes[offset + 2])
                    | (UInt32(bytes[offset + 3]) << 8)
                    | (UInt32(bytes[offset + 4]) << 16)
                    | (UInt32(bytes[offset + 5]) << 24)
            return Int(Int32(bitPattern: raw))
        default:
            return nil
        }
    }

    private static func readStringValue(from data: Data, at offset: Int) -> String? {
        let bytes = [UInt8](data)
        guard offset + 6 <= bytes.count else { return nil }
        let vt = UInt16(bytes[offset]) | (UInt16(bytes[offset + 1]) << 8)
        guard vt == VT_LPSTR else { return nil }

        let length = Int(
            UInt32(bytes[offset + 2])
            | (UInt32(bytes[offset + 3]) << 8)
            | (UInt32(bytes[offset + 4]) << 16)
            | (UInt32(bytes[offset + 5]) << 24)
        )

        let strStart = offset + 6
        guard length > 0, strStart + length <= bytes.count else { return nil }

        let strBytes = Array(bytes[strStart..<strStart + length])
        // VT_LPSTR is a null-terminated ANSI string
        let nullTerminated = strBytes.prefix(while: { $0 != 0 })
        return String(bytes: nullTerminated, encoding: .windowsCP1252)
            ?? String(bytes: nullTerminated, encoding: .utf8)
    }

    private static func readFileTimeValue(from data: Data, at offset: Int) -> Date? {
        let bytes = [UInt8](data)
        guard offset + 10 <= bytes.count else { return nil }
        let vt = UInt16(bytes[offset]) | (UInt16(bytes[offset + 1]) << 8)
        guard vt == VT_FILETIME else { return nil }

        // FILETIME: 64-bit unsigned, 100-ns intervals since 1601-01-01 UTC
        let raw = UInt64(bytes[offset + 2])
                | (UInt64(bytes[offset + 3]) << 8)
                | (UInt64(bytes[offset + 4]) << 16)
                | (UInt64(bytes[offset + 5]) << 24)
                | (UInt64(bytes[offset + 6]) << 32)
                | (UInt64(bytes[offset + 7]) << 40)
                | (UInt64(bytes[offset + 8]) << 48)
                | (UInt64(bytes[offset + 9]) << 56)

        guard raw > 0 else { return nil }

        // Seconds between Windows epoch (1601-01-01) and Unix epoch (1970-01-01)
        let windowsToUnixSeconds: Double = 11_644_473_600
        let seconds = Double(raw) / 10_000_000 - windowsToUnixSeconds
        return Date(timeIntervalSince1970: seconds)
    }
}
