"""產生 FileInspector 測試用的範例檔（正常、各種毀損、IRM / 密碼保護的 OLE 容器）。

用法：python tests/make_fixtures.py
輸出：tests/fixtures/*.xlsx|docx|xls 與 tests/fixtures/expected.json
"""
import io
import json
import os
import random
import struct
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "fixtures")

LABEL_ID = "6f1c2c49-6b5a-4d1e-9a51-7a3c1e0d4b21"
SITE_ID = "0e6c3c7a-1b2d-4c5e-8f90-a1b2c3d4e5f6"

CT_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def content_types(overrides, extra_defaults=()):
    xml = CT_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    xml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    xml += '<Default Extension="xml" ContentType="application/xml"/>'
    for ext, ct in extra_defaults:
        xml += '<Default Extension="%s" ContentType="%s"/>' % (ext, ct)
    for part, ct in overrides:
        xml += '<Override PartName="/%s" ContentType="%s"/>' % (part, ct)
    return xml + "</Types>"


def rels(items):
    xml = CT_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for rid, typ, target in items:
        xml += '<Relationship Id="%s" Type="%s" Target="%s"/>' % (rid, typ, target)
    return xml + "</Relationships>"


R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
SHEET_TEXT = "UNLOCKED_MARKER_TEXT_0123456789"


def sheet_xml(text=SHEET_TEXT):
    return (CT_HEAD + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>%s</t></is></c></row></sheetData></worksheet>' % text)


def xlsx_parts(sheet=None, sheets=1, label=False, missing_sheet=False):
    workbook = (CT_HEAD + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>')
    wb_rels = []
    overrides = [("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")]
    parts = {}
    for i in range(1, sheets + 1):
        workbook += '<sheet name="Sheet%d" sheetId="%d" r:id="rId%d"/>' % (i, i, i)
        wb_rels.append(("rId%d" % i, R + "worksheet", "worksheets/sheet%d.xml" % i))
        overrides.append(("xl/worksheets/sheet%d.xml" % i,
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))
        if not (missing_sheet and i == sheets):
            parts["xl/worksheets/sheet%d.xml" % i] = sheet if (sheet is not None and i == 1) else sheet_xml()
    workbook += "</sheets></workbook>"

    root_rels = [("rId1", R + "officeDocument", "xl/workbook.xml")]
    if label:
        root_rels.append(("rId2", "http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels",
                          "docMetadata/LabelInfo.xml"))
        root_rels.append(("rId3", R + "custom-properties", "docProps/custom.xml"))
        overrides.append(("docMetadata/LabelInfo.xml", "application/vnd.ms-office.classificationlabels+xml"))
        overrides.append(("docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml"))
        parts["docMetadata/LabelInfo.xml"] = (
            CT_HEAD + '<clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata">'
            '<clbl:label id="{%s}" enabled="1" method="Privileged" siteId="{%s}" removed="0"/></clbl:labelList>'
            % (LABEL_ID, SITE_ID))
        props = ""
        pid = 2
        for field, value in (("Enabled", "true"), ("Method", "Privileged"), ("Name", "Public-Edit"),
                             ("SiteId", SITE_ID), ("ContentBits", "0")):
            props += ('<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="MSIP_Label_%s_%s">'
                      '<vt:lpwstr>%s</vt:lpwstr></property>' % (pid, LABEL_ID, field, value))
            pid += 1
        parts["docProps/custom.xml"] = (
            CT_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">%s</Properties>' % props)

    ordered = {"[Content_Types].xml": content_types(overrides), "_rels/.rels": rels(root_rels),
               "xl/workbook.xml": workbook, "xl/_rels/workbook.xml.rels": rels(wb_rels)}
    ordered.update(parts)
    return ordered


def docx_parts():
    return {
        "[Content_Types].xml": content_types(
            [("word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]),
        "_rels/.rels": rels([("rId1", R + "officeDocument", "word/document.xml")]),
        "word/document.xml": (CT_HEAD + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                              '<w:body><w:p><w:r><w:t>UnLocked</w:t></w:r></w:p></w:body></w:document>'),
    }


def pptx_parts(missing_slide=False):
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    parts = {
        "[Content_Types].xml": content_types([
            ("ppt/presentation.xml", "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"),
            ("ppt/slides/slide1.xml", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")]),
        "_rels/.rels": rels([("rId1", R + "officeDocument", "ppt/presentation.xml")]),
        "ppt/presentation.xml": (CT_HEAD + '<p:presentation xmlns:p="%s" '
                                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                                 '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>' % P),
        "ppt/_rels/presentation.xml.rels": rels([("rId1", R + "slide", "slides/slide1.xml")]),
        "ppt/slides/slide1.xml": (CT_HEAD + '<p:sld xmlns:p="%s"><p:cSld><p:spTree/></p:cSld></p:sld>' % P),
    }
    if missing_slide:
        del parts["ppt/slides/slide1.xml"]
    return parts


def write_zip(name, parts, stored=()):
    path = os.path.join(OUT, name)
    with zipfile.ZipFile(path, "w") as z:
        for part, data in parts.items():
            info = zipfile.ZipInfo(part, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_STORED if part in stored else zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode("utf-8") if isinstance(data, str) else data)
    return path


# ---------------------------------------------------------------------------
#  最小 OLE 複合檔 (CFB v3, 512 位元組磁區，所有資料流 >= 4096 位元組，不使用迷你資料流)
# ---------------------------------------------------------------------------
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


def write_cfb(name, tree):
    """tree: [(name, bytes) | (name, [children...])] 為根目錄下的項目。"""
    entries = []  # dict(name, type, data, children)

    def add(node_name, payload):
        idx = len(entries)
        e = {"name": node_name, "children": []}
        entries.append(e)
        if isinstance(payload, list):
            e["type"] = 1
            e["data"] = b""
            e["children"] = [add(n, p) for n, p in payload]
        else:
            e["type"] = 2
            e["data"] = payload
        return idx

    entries.append({"name": "Root Entry", "type": 5, "data": b"", "children": []})
    entries[0]["children"] = [add(n, p) for n, p in tree]

    ss = 512
    dir_sectors = (len(entries) * 128 + ss - 1) // ss
    data_start = 1 + dir_sectors
    next_sector = data_start
    fat = [FATSECT] + [ENDOFCHAIN if i == dir_sectors - 1 else 1 + i + 1 for i in range(dir_sectors)]
    body = b""
    for e in entries:
        if e["type"] != 2:
            e["start"] = ENDOFCHAIN
            continue
        assert len(e["data"]) >= 4096
        n = (len(e["data"]) + ss - 1) // ss
        e["start"] = next_sector
        for k in range(n):
            fat.append(ENDOFCHAIN if k == n - 1 else next_sector + k + 1)
        next_sector += n
        body += e["data"] + b"\0" * (n * ss - len(e["data"]))
    assert len(fat) <= ss // 4
    fat += [FREESECT] * (ss // 4 - len(fat))

    # 兄弟節點以右側鏈結串起（全部為黑色節點，結構合法）
    for e in entries:
        e.setdefault("left", NOSTREAM)
        e.setdefault("right", NOSTREAM)
        kids = e["children"]
        e["child"] = kids[0] if kids else NOSTREAM
        for a, b in zip(kids, kids[1:]):
            entries[a]["right"] = b

    d = b""
    for e in entries:
        nm = e["name"].encode("utf-16-le")
        rec = nm + b"\0" * (64 - len(nm))
        rec += struct.pack("<HBB", len(nm) + 2, e["type"], 1)
        rec += struct.pack("<III", e["left"], e["right"], e["child"])
        rec += b"\0" * 16 + b"\0" * 4 + b"\0" * 16
        rec += struct.pack("<IQ", e["start"], len(e["data"]))
        assert len(rec) == 128
        d += rec
    d += b"\0" * (dir_sectors * ss - len(d))

    header = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\0" * 16
    header += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
    header += struct.pack("<IIIIIIIII", 0, 1, 1, 0, 4096, ENDOFCHAIN, 0, ENDOFCHAIN, 0)
    header += struct.pack("<I", 0) + struct.pack("<I", FREESECT) * 108
    assert len(header) == 512
    data = header + struct.pack("<%dI" % (ss // 4), *fat) + d + body
    path = os.path.join(OUT, name)
    with open(path, "wb") as f:
        f.write(data)
    return path


def main():
    os.makedirs(OUT, exist_ok=True)
    expected = {}

    write_zip("valid.xlsx", xlsx_parts())
    expected["valid.xlsx"] = {"Container": "OOXML", "Status": "OK", "DocKind": "Excel", "MainPart": "xl/workbook.xml",
                              "IrmProtected": False}

    write_zip("valid.docx", docx_parts())
    expected["valid.docx"] = {"Container": "OOXML", "Status": "OK", "DocKind": "Word"}

    write_zip("valid.pptx", pptx_parts())
    expected["valid.pptx"] = {"Container": "OOXML", "Status": "OK", "DocKind": "PowerPoint",
                              "MainPart": "ppt/presentation.xml"}

    write_zip("missingslide.pptx", pptx_parts(missing_slide=True))
    expected["missingslide.pptx"] = {"Status": "CORRUPT", "DocKind": "PowerPoint", "IssueContains": "slide1.xml"}

    write_zip("labeled.xlsx", xlsx_parts(label=True))
    expected["labeled.xlsx"] = {"Status": "OK", "LabelId": LABEL_ID, "LabelSiteId": SITE_ID,
                                "LabelName": "Public-Edit", "LabelMethod": "Privileged"}

    p = write_zip("_tmp.xlsx", xlsx_parts(sheets=2))
    raw = open(p, "rb").read()
    os.remove(p)
    with open(os.path.join(OUT, "truncated.xlsx"), "wb") as f:
        f.write(raw[: int(len(raw) * 0.6)])
    expected["truncated.xlsx"] = {"Container": "OOXML", "Status": "CORRUPT", "IssueContains": "中央目錄"}

    write_zip("badxml.xlsx", xlsx_parts(sheet=CT_HEAD + "<worksheet><sheetData><row></sheetData></worksheet>"))
    expected["badxml.xlsx"] = {"Status": "CORRUPT", "IssueContains": "XML 格式錯誤"}

    write_zip("badchar.xlsx", xlsx_parts(sheet=sheet_xml("bad\x01char")))
    expected["badchar.xlsx"] = {"Status": "CORRUPT", "IssueContains": "sheet1.xml"}

    write_zip("missingpart.xlsx", xlsx_parts(sheets=2, missing_sheet=True))
    expected["missingpart.xlsx"] = {"Status": "CORRUPT", "IssueContains": "sheet2.xml"}

    parts = xlsx_parts()
    del parts["[Content_Types].xml"]
    write_zip("nocontenttypes.xlsx", parts)
    expected["nocontenttypes.xlsx"] = {"Status": "CORRUPT", "IssueContains": "[Content_Types].xml"}

    p = write_zip("crc.xlsx", xlsx_parts(), stored=("xl/worksheets/sheet1.xml",))
    raw = bytearray(open(p, "rb").read())
    pos = raw.find(SHEET_TEXT.encode())
    assert pos > 0
    raw[pos] = ord("X")
    open(p, "wb").write(bytes(raw))
    expected["crc.xlsx"] = {"Status": "CORRUPT", "IssueContains": "CRC"}

    open(os.path.join(OUT, "empty.xlsx"), "wb").close()
    expected["empty.xlsx"] = {"Container": "EMPTY", "Status": "CORRUPT"}

    rnd = random.Random(20260115)
    with open(os.path.join(OUT, "unknown.xlsx"), "wb") as f:
        f.write(b"DRM!" + bytes(rnd.randrange(256) for _ in range(4000)))
    expected["unknown.xlsx"] = {"Container": "UNKNOWN", "Status": "UNKNOWN"}

    blob = bytes(rnd.randrange(256) for _ in range(5000))
    write_cfb("irm.xlsx", [("EncryptedPackage", blob),
                           ("\x06DataSpaces", [("DataSpaceMap", b"\x08" * 4096),
                                               ("DataSpaceInfo", [("DRMEncryptedDataSpace", b"\x01" * 4096)])])])
    expected["irm.xlsx"] = {"Container": "CFB", "Status": "PROTECTED", "IrmProtected": True, "PasswordEncrypted": False}

    write_cfb("password.xlsx", [("EncryptionInfo", b"\x04" * 4096), ("EncryptedPackage", blob),
                                ("\x06DataSpaces", [("DataSpaceInfo", [("StrongEncryptionDataSpace", b"\x02" * 4096)])])])
    expected["password.xlsx"] = {"Container": "CFB", "Status": "PASSWORD", "PasswordEncrypted": True, "IrmProtected": False}

    p = write_cfb("legacy.xls", [("Workbook", b"\x09\x08" * 4096)])
    expected["legacy.xls"] = {"Container": "CFB", "Status": "OK", "DocKind": "Excel"}
    write_cfb("legacy.doc", [("WordDocument", b"\xec\xa5" * 2048), ("1Table", b"\0" * 4096)])
    expected["legacy.doc"] = {"Container": "CFB", "Status": "OK", "DocKind": "Word"}

    write_cfb("legacy.ppt", [("PowerPoint Document", b"\0\x0f" * 2048), ("Current User", b"\0" * 4096)])
    expected["legacy.ppt"] = {"Container": "CFB", "Status": "OK", "DocKind": "PowerPoint"}

    raw = open(p, "rb").read()
    with open(os.path.join(OUT, "legacy_truncated.xls"), "wb") as f:
        f.write(raw[: 512 * 8])
    expected["legacy_truncated.xls"] = {"Container": "CFB", "Status": "CORRUPT"}

    with open(os.path.join(OUT, "expected.json"), "w", encoding="utf-8") as f:
        json.dump(expected, f, ensure_ascii=False, indent=2)
    print("wrote %d fixtures to %s" % (len(expected), OUT))


if __name__ == "__main__":
    main()
