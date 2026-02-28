#!/usr/bin/env python3
"""
Generate DocReader integration test fixtures.

OOXML (.docx, .xlsx, .pptx) are built from raw ZIP+XML for precise metadata.
Legacy (.xls) uses xlwt.  Legacy (.doc, .ppt) use a minimal OLE2 writer.
"""

import io, os, struct, zipfile

FIXTURES_DIR = os.path.join(
    os.path.dirname(__file__),
    "..", "Tests", "DocReaderIntegrationTests", "Fixtures"
)

# ── OLE2 Compound File writer ────────────────────────────────────────────────

FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
NOSTREAM   = 0xFFFFFFFF
SECTOR     = 512


def _dir_entry(name: str, obj_type: int, start: int, size: int,
               child=NOSTREAM, left=NOSTREAM, right=NOSTREAM) -> bytes:
    """Build a 128-byte OLE2 directory entry."""
    name_u = name.encode("utf-16-le")
    name_len = len(name_u) + 2          # include null terminator
    name_field = (name_u + b"\x00\x00").ljust(64, b"\x00")
    e  = name_field                     # 64
    e += struct.pack("<H", name_len)    # 2
    e += struct.pack("<B", obj_type)    # 1  (2=stream, 5=root)
    e += struct.pack("<B", 1)           # 1  color=black
    e += struct.pack("<I", left)        # 4
    e += struct.pack("<I", right)       # 4
    e += struct.pack("<I", child)       # 4
    e += b"\x00" * 16                   # CLSID
    e += struct.pack("<I", 0)           # state bits
    e += struct.pack("<Q", 0)           # created FILETIME
    e += struct.pack("<Q", 0)           # modified FILETIME
    e += struct.pack("<I", start)       # start sector
    e += struct.pack("<I", size)        # size low
    e += struct.pack("<I", 0)           # size high (v3)
    assert len(e) == 128
    return e


def build_ole2(streams: list) -> bytes:
    """
    Build a minimal, valid OLE2 CFB v3 file.
    streams: [(name, data_bytes), ...]   Max 3 streams (fits in 1 dir sector).

    Sector layout: 0=FAT, 1=DIR, 2=MiniFAT (empty), 3..=stream data.

    OLEKit (Swift) unconditionally calls loadMiniFAT() during OLEFile.init,
    so a real mini-FAT sector must exist even when no mini-stream data is
    stored.  All stream data must be >= 4096 bytes (miniStreamCutoffSize) so
    OLEKit routes them through the regular FAT, not the (empty) mini-stream.
    """
    MINIFAT_SECTOR = 2
    DATA_START     = 3

    # Assign sectors for stream data (starting after mini-FAT sector)
    sector_starts = []
    cur = DATA_START
    for _, data in streams:
        sector_starts.append(cur)
        cur += max(1, -(-len(data) // SECTOR))   # ceil division

    # FAT
    fat = [FREESECT] * 128
    fat[0] = FATSECT                        # sector 0 = FAT
    fat[1] = ENDOFCHAIN                     # sector 1 = DIR (end of chain)
    fat[MINIFAT_SECTOR] = ENDOFCHAIN        # sector 2 = MiniFAT (end of chain)
    for i, (_, data) in enumerate(streams):
        n = max(1, -(-len(data) // SECTOR))
        s = sector_starts[i]
        for j in range(n):
            fat[s + j] = (s + j + 1) if j < n - 1 else ENDOFCHAIN
    fat_sector = struct.pack("<128I", *fat)

    # Mini-FAT sector: all FREESECT (no mini-stream contents)
    minifat_sector = struct.pack("<128I", *([FREESECT] * 128))

    # Directory (1 sector = 4 entries)
    # Entry 0: Root  Entry 1..N: streams (right-sibling chain)
    n = len(streams)
    root = _dir_entry("Root Entry", 5, ENDOFCHAIN, 0,
                      child=1 if n else NOSTREAM)
    entries = [root]
    for i, (name, data) in enumerate(streams):
        right = (i + 2) if i + 1 < n else NOSTREAM
        entries.append(_dir_entry(name, 2, sector_starts[i], len(data),
                                  right=right))
    while len(entries) < 4:
        entries.append(b"\x00" * 128)
    dir_sector = b"".join(entries)
    assert len(dir_sector) == SECTOR

    # Header (512 bytes)
    hdr  = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"   # magic       8
    hdr += b"\x00" * 16                             # CLSID      16
    hdr += struct.pack("<HH", 0x003E, 0x0003)       # ver        4
    hdr += struct.pack("<H",  0xFFFE)               # byte order 2
    hdr += struct.pack("<HH", 0x0009, 0x0006)       # sector shifts 4
    hdr += b"\x00" * 6                              # reserved   6
    hdr += struct.pack("<I", 0)                     # dir sectors 4
    hdr += struct.pack("<I", 1)                     # FAT sectors (1) 4
    hdr += struct.pack("<I", 1)                     # first dir sector = 1 4
    hdr += struct.pack("<I", 0)                     # txn sig    4
    hdr += struct.pack("<I", 0x1000)                # mini cutoff 4
    hdr += struct.pack("<I", MINIFAT_SECTOR)        # mini FAT first sector 4
    hdr += struct.pack("<I", 1)                     # mini FAT sector count 4
    hdr += struct.pack("<I", ENDOFCHAIN)            # DIFAT start 4
    hdr += struct.pack("<I", 0)                     # DIFAT cnt  4
    hdr += struct.pack("<I", 0)                     # DIFAT[0]=FAT sector 0
    hdr += struct.pack("<I", FREESECT) * 108        # DIFAT[1..108]
    assert len(hdr) == SECTOR

    # Stream data (padded to sector boundary)
    body = b""
    for _, data in streams:
        rem = len(data) % SECTOR
        body += data + (b"\x00" * (SECTOR - rem) if rem else b"")

    return hdr + fat_sector + dir_sector + minifat_sector + body


# ── OLE2 SummaryInformation property set ────────────────────────────────────

# The OLEPropertySetReader expects no padding between type-code and value:
#   VT_I4   → 2-byte type + 4-byte int
#   VT_LPSTR → 2-byte type + 4-byte count + count bytes (ANSI + null)

VT_I4    = 0x0003
VT_LPSTR = 0x001E

# FMTID {F29F85E0-4FF9-1068-AB91-08002B27B3D9} (SummaryInformation)
_FMTID_SUMMARY = bytes([
    0xE0, 0x85, 0x9F, 0xF2,   # Data1 LE
    0xF9, 0x4F,               # Data2 LE
    0x68, 0x10,               # Data3 LE
    0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9,  # Data4 BE
])

SECTION_OFFSET = 48   # = 28-byte header + 20-byte list entry


def _prop_i4(pid: int, value: int):
    return pid, struct.pack("<H", VT_I4) + struct.pack("<i", value)


def _prop_str(pid: int, text: str):
    data = text.encode("latin-1", errors="replace") + b"\x00"
    if len(data) % 4:
        data += b"\x00" * (4 - len(data) % 4)
    return pid, struct.pack("<H", VT_LPSTR) + struct.pack("<I", len(data)) + data


def make_summary_stream(page_count: int,
                        title: str = "", author: str = "") -> bytes:
    props = []
    if title:
        props.append(_prop_str(0x02, title))
    if author:
        props.append(_prop_str(0x04, author))
    props.append(_prop_i4(0x0E, page_count))   # PIDSI_PAGECOUNT

    # Build PropertySet section
    id_offset_size = 8 * len(props)
    header_size = 8 + id_offset_size            # Size(4) + NumProps(4) + pairs

    offsets = []
    off = header_size
    for _, val in props:
        offsets.append(off)
        off += len(val)
        off = (off + 3) & ~3                    # 4-byte align

    prop_set = struct.pack("<II", off, len(props))
    for (pid, _), o in zip(props, offsets):
        prop_set += struct.pack("<II", pid, o)
    for i, (_, val) in enumerate(props):
        prop_set += val
        if i < len(props) - 1:
            pad = (4 - len(val) % 4) % 4
            prop_set += b"\x00" * pad

    # Stream header (48 bytes)
    stream  = struct.pack("<HH", 0xFFFE, 0x0000)   # ByteOrder, Version
    stream += struct.pack("<I",  0x00020006)         # SystemIdentifier
    stream += b"\x00" * 16                           # CLSID
    stream += struct.pack("<I",  1)                  # NumPropertySets
    stream += _FMTID_SUMMARY                         # FMTID (16)
    stream += struct.pack("<I",  SECTION_OFFSET)     # offset to PropertySet
    assert len(stream) == SECTION_OFFSET

    return stream + prop_set


# ── Minimal OOXML builders ───────────────────────────────────────────────────

_CONTENT_TYPES_DOCX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/app.xml"
    ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

_RELS_DOCX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="word/document.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
  <Relationship Id="rId2" Target="docProps/core.xml"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>
  <Relationship Id="rId3" Target="docProps/app.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"/>
</Relationships>"""

_DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"""


def _docx_document(page_count: int) -> str:
    """Word document.xml with one paragraph per page (explicit page breaks)."""
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    paras = []
    for i in range(page_count):
        if i > 0:
            paras.append("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>")
        paras.append(f"<w:p><w:r><w:t>Page {i+1} content.</w:t></w:r></w:p>")
    body = "\n    ".join(paras)
    # US Letter: 12240 × 15840 twips  (612pt × 792pt)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {ns}>
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>"""


def _app_xml(pages: int, app: str = "DocReader Fixture Generator") -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>{app}</Application>
  <Pages>{pages}</Pages>
</Properties>"""


def _core_xml(title: str, creator: str = "DocReader Tests") -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{title}</dc:title>
  <dc:creator>{creator}</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-02-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""


def make_docx(filename: str, page_count: int, title: str) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",  _CONTENT_TYPES_DOCX)
        zf.writestr("_rels/.rels",           _RELS_DOCX)
        zf.writestr("word/_rels/document.xml.rels", _DOC_RELS)
        zf.writestr("word/document.xml",    _docx_document(page_count))
        zf.writestr("docProps/app.xml",     _app_xml(page_count))
        zf.writestr("docProps/core.xml",    _core_xml(title))
    path = os.path.join(FIXTURES_DIR, filename)
    with open(path, "wb") as f:
        f.write(buf.getvalue())
    print(f"✓ {filename}  ({len(buf.getvalue())} bytes)")


# ── XLSX ──────────────────────────────────────────────────────────────────────

_CONTENT_TYPES_XLSX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

_RELS_XLSX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="xl/workbook.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
  <Relationship Id="rId2" Target="docProps/core.xml"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>
</Relationships>"""

_WB_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="worksheets/sheet1.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
  <Relationship Id="rId2" Target="worksheets/sheet2.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
  <Relationship Id="rId3" Target="worksheets/sheet3.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
</Relationships>"""

_WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
    <sheet name="Sheet3" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"""


def _sheet_xml(name: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>{name}</t></is></c></row>
  </sheetData>
  <pageSetup paperSize="9"/>
</worksheet>"""


def make_xlsx_3sheet() -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",        _CONTENT_TYPES_XLSX)
        zf.writestr("_rels/.rels",                _RELS_XLSX)
        zf.writestr("xl/_rels/workbook.xml.rels", _WB_RELS)
        zf.writestr("xl/workbook.xml",            _WORKBOOK)
        for i in range(1, 4):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(f"Sheet{i}"))
        zf.writestr("docProps/core.xml",
                    _core_xml("DocReader XLSX Fixture"))
    path = os.path.join(FIXTURES_DIR, "excel_3sheet.xlsx")
    with open(path, "wb") as f:
        f.write(buf.getvalue())
    print(f"✓ excel_3sheet.xlsx  ({len(buf.getvalue())} bytes)")


# ── PPTX ──────────────────────────────────────────────────────────────────────

_CONTENT_TYPES_PPTX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/docProps/core.xml"
    ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

_RELS_PPTX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="ppt/presentation.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
  <Relationship Id="rId2" Target="docProps/core.xml"
    Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>
</Relationships>"""

_PPT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"""


def _presentation_xml(slide_count: int) -> str:
    # Widescreen 16:9: 9144000 × 5143500 EMU  →  720 × 405 pt
    sld_ids = "\n    ".join(
        f'<p:sldId id="{256 + i}" r:id="rId{i+1}"/>'
        for i in range(slide_count)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="9144000" cy="5143500"/>
  <p:sldIdLst>
    {sld_ids}
  </p:sldIdLst>
</p:presentation>"""


def make_pptx_5slide() -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",          _CONTENT_TYPES_PPTX)
        zf.writestr("_rels/.rels",                  _RELS_PPTX)
        zf.writestr("ppt/_rels/presentation.xml.rels", _PPT_RELS)
        zf.writestr("ppt/presentation.xml",         _presentation_xml(5))
        zf.writestr("docProps/core.xml",
                    _core_xml("DocReader PPTX Fixture"))
    path = os.path.join(FIXTURES_DIR, "ppt_5slide.pptx")
    with open(path, "wb") as f:
        f.write(buf.getvalue())
    print(f"✓ ppt_5slide.pptx  ({len(buf.getvalue())} bytes)")


# ── Legacy XLS ────────────────────────────────────────────────────────────────

def make_xls_3sheet() -> None:
    import xlwt
    wb = xlwt.Workbook()
    for i in range(1, 4):
        ws = wb.add_sheet(f"Sheet{i}")
        ws.write(0, 0, f"DocReader legacy test sheet {i}")
    path = os.path.join(FIXTURES_DIR, "excel_legacy_3sheet.xls")
    wb.save(path)

    # xlwt writes firstMiniFATSector = ENDOFCHAIN (0xFFFFFFFE) when there is
    # no mini-FAT.  OLEKit (Swift) throws OLEError.invalidEmptyStream when
    # sectorID == ENDOFCHAIN && expectedStreamSize == 0.  Patch to FREESECT
    # (0xFFFFFFFF) which passes the guard while still signalling "no data".
    # Header offset 60 = _sectMiniFatStart (4 bytes, little-endian).
    with open(path, "r+b") as f:
        f.seek(60)
        current = struct.unpack("<I", f.read(4))[0]
        if current == 0xFFFFFFFE:   # ENDOFCHAIN → patch to FREESECT
            f.seek(60)
            f.write(struct.pack("<I", 0xFFFFFFFF))

    print(f"✓ excel_legacy_3sheet.xls  ({os.path.getsize(path)} bytes)")


# ── Legacy DOC ────────────────────────────────────────────────────────────────

def make_doc_10page() -> None:
    summary = make_summary_stream(
        page_count=10, title="Legacy Word Test", author="DocReader Tests"
    )
    # Pad to 4096 bytes so OLEKit routes via FAT (not mini-stream).
    # OLEKit rejects streams < miniStreamCutoffSize (0x1000) when no
    # mini-stream infrastructure exists in the file.
    summary = summary.ljust(4096, b"\x00")
    data = build_ole2([("\x05SummaryInformation", summary)])
    path = os.path.join(FIXTURES_DIR, "word_legacy_10page.doc")
    with open(path, "wb") as f:
        f.write(data)
    print(f"✓ word_legacy_10page.doc  ({len(data)} bytes)")


# ── Legacy PPT ────────────────────────────────────────────────────────────────

def make_ppt_5slide() -> None:
    """5 SlideContainer atoms (type 0x03E8) with length = 0."""
    # PPT record: version+instance (2) | type (2) | length (4)
    slide_stream = b""
    for _ in range(5):
        slide_stream += struct.pack("<HHI",
            0x000F,   # version=15 (container), instance=0
            0x03E8,   # SlideContainer record type
            0,        # body length = 0
        )
    # Pad to 4096 bytes so OLEKit routes via FAT (not mini-stream).
    # countSlideContainerAtoms only counts type 0x03E8; zero-padded bytes
    # produce type=0 records that are skipped harmlessly.
    slide_stream = slide_stream.ljust(4096, b"\x00")
    data = build_ole2([("PowerPoint Document", slide_stream)])
    path = os.path.join(FIXTURES_DIR, "ppt_legacy_5slide.ppt")
    with open(path, "wb") as f:
        f.write(data)
    print(f"✓ ppt_legacy_5slide.ppt  ({len(data)} bytes)")


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    os.makedirs(FIXTURES_DIR, exist_ok=True)

    print("Generating OOXML fixtures…")
    make_docx("word_1page.docx",   page_count=1,  title="One-Page Test Document")
    make_docx("word_10page.docx",  page_count=10, title="Ten-Page Test Document")
    make_xlsx_3sheet()
    make_pptx_5slide()

    print("\nGenerating legacy fixtures…")
    make_xls_3sheet()
    make_doc_10page()
    make_ppt_5slide()

    print("\nDone — all 7 fixtures written to:")
    print(f"  {os.path.abspath(FIXTURES_DIR)}")
