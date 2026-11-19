#!/usr/bin/env python3
"""Zero-dependency OpenXML (.xlsx and .pptx) builders using Python's standard library."""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any, List, Sequence, Tuple
from xml.sax.saxutils import escape


def _col_letter(col_idx_1based: int) -> str:
    out = ""
    n = col_idx_1based
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def write_multi_tab_xlsx(
    out_path: Path,
    sheets: Sequence[Tuple[str, Sequence[Sequence[Any]]]],
) -> None:
    """Write a valid OpenXML .xlsx workbook supporting numbers, text, and '=...' formulas."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    safe_names = [name[:31] for name, _ in sheets]

    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for i in range(1, len(sheets) + 1):
        content_types.append(
            f'  <Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    content_types.append("</Types>")

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
        "</Relationships>"
    )

    wb_sheets = []
    wb_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for i, sname in enumerate(safe_names, 1):
        wb_sheets.append(f'    <sheet name="{escape(sname)}" sheetId="{i}" r:id="rId{i}"/>')
        wb_rels.append(
            f'  <Relationship Id="rId{i}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i}.xml"/>'
        )
    wb_rels.append(
        f'  <Relationship Id="rId{len(sheets) + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    wb_rels.append("</Relationships>")

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        "  <sheets>\n"
        + "\n".join(wb_sheets)
        + "\n  </sheets>\n</workbook>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        '  <fonts count="2">\n'
        '    <font><sz val="10"/><color rgb="FF20272C"/><name val="Calibri"/></font>\n'
        '    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>\n'
        "  </fonts>\n"
        '  <fills count="3">\n'
        '    <fill><patternFill patternType="none"/></fill>\n'
        '    <fill><patternFill patternType="gray125"/></fill>\n'
        '    <fill><patternFill patternType="solid"><fgColor rgb="FF173B55"/></patternFill></fill>\n'
        "  </fills>\n"
        '  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n'
        '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'
        '  <cellXfs count="2">\n'
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n'
        '    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>\n'
        "  </cellXfs>\n"
        "</styleSheet>"
    )

    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "\n".join(content_types))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", "\n".join(wb_rels))
        zf.writestr("xl/styles.xml", styles_xml)

        for s_idx, (_, rows) in enumerate(sheets, 1):
            row_xml: List[str] = []
            for r_idx, row in enumerate(rows, 1):
                style_attr = ' s="1"' if r_idx == 1 else ""
                cells: List[str] = []
                for c_idx, val in enumerate(row, 1):
                    ref = f"{_col_letter(c_idx)}{r_idx}"
                    if val is None or val == "":
                        continue
                    if isinstance(val, bool):
                        cells.append(f'<c r="{ref}" t="b"{style_attr}><v>{1 if val else 0}</v></c>')
                    elif isinstance(val, (int, float)):
                        cells.append(f'<c r="{ref}"{style_attr}><v>{val}</v></c>')
                    else:
                        sval = str(val)
                        if sval.startswith("="):
                            formula = escape(sval[1:])
                            cells.append(f'<c r="{ref}"{style_attr}><f>{formula}</f></c>')
                        else:
                            cells.append(
                                f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{escape(sval)}</t></is></c>'
                            )
                row_xml.append(f'    <row r="{r_idx}">{"".join(cells)}</row>')
            sheet_doc = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
                "  <sheetData>\n"
                + "\n".join(row_xml)
                + "\n  </sheetData>\n</worksheet>"
            )
            zf.writestr(f"xl/worksheets/sheet{s_idx}.xml", sheet_doc)


def write_screenshot_deck_pptx(
    out_path: Path,
    slides: Sequence[Tuple[str, str, Sequence[str]]],
) -> None:
    """Write a valid minimal OpenXML .pptx presentation deck for campaign proof-of-launch tearsheets."""
    out_path.parent.mkdir(parents=True, exist_ok=True)
    n_slides = len(slides)

    ct_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
        '  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    ]
    for i in range(1, n_slides + 1):
        ct_lines.append(
            f'  <Override PartName="/ppt/slides/slide{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )
    ct_lines.append("</Types>")

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>\n'
        "</Relationships>"
    )

    pres_rels = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>',
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
    ]
    sld_ids = []
    for i in range(1, n_slides + 1):
        rid = f"rId{i + 2}"
        pres_rels.append(
            f'  <Relationship Id="{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            f'Target="slides/slide{i}.xml"/>'
        )
        sld_ids.append(f'    <p:sldId id="{255 + i}" r:id="{rid}"/>')
    pres_rels.append("</Relationships>")

    presentation_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\n'
        '  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>\n'
        "  <p:sldIdLst>\n"
        + "\n".join(sld_ids)
        + '\n  </p:sldIdLst>\n  <p:sldSz cx="12192000" cy="6858000"/>\n  <p:notesSz cx="6858000" cy="9144000"/>\n'
        "</p:presentation>"
    )

    master_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>\n'
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>\n'
        "</Relationships>"
    )
    master_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\n'
        '  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr/></p:spTree></p:cSld>\n'
        '  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
        'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>\n'
        '  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>\n'
        "</p:sldMaster>"
    )
    layout_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>\n'
        "</Relationships>"
    )
    layout_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">\n'
        '  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr/></p:spTree></p:cSld>\n'
        "</p:sldLayout>"
    )
    theme_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="AdOpsTheme">\n'
        '  <a:themeElements>\n'
        '    <a:clrScheme name="AdOps"><a:dk1><a:srgbClr val="20272C"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="173B55"/></a:dk2><a:lt2><a:srgbClr val="F4F6F8"/></a:lt2>'
        '<a:accent1><a:srgbClr val="173B55"/></a:accent1><a:accent2><a:srgbClr val="2B6CB0"/></a:accent2>'
        '<a:accent3><a:srgbClr val="2F855A"/></a:accent3><a:accent4><a:srgbClr val="D69E2E"/></a:accent4>'
        '<a:accent5><a:srgbClr val="C53030"/></a:accent5><a:accent6><a:srgbClr val="718096"/></a:accent6>'
        '<a:hlink><a:srgbClr val="173B55"/></a:hlink><a:folHlink><a:srgbClr val="4A5568"/></a:folHlink></a:clrScheme>\n'
        '    <a:fontScheme name="AdOpsFont"><a:majorFont><a:latin typeface="Arial"/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme>\n'
        '    <a:fmtScheme name="AdOpsFmt"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
        '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
        '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
        '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>\n'
        "  </a:themeElements>\n</a:theme>"
    )
    slide_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>\n'
        "</Relationships>"
    )

    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", "\n".join(ct_lines))
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("ppt/presentation.xml", presentation_xml)
        zf.writestr("ppt/_rels/presentation.xml.rels", "\n".join(pres_rels))
        zf.writestr("ppt/slideMasters/slideMaster1.xml", master_xml)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", layout_xml)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout_rels)
        zf.writestr("ppt/theme/theme1.xml", theme_xml)

        for i, (title, placement_banner, meta_lines) in enumerate(slides, 1):
            paras = "".join(
                f'<a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>{escape(line)}</a:t></a:r></a:p>'
                for line in meta_lines
            )
            slide_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\n'
                "  <p:cSld><p:spTree>\n"
                '    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>\n'
                '    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Header"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>\n'
                '      <p:spPr><a:xfrm><a:off x="457200" y="274320"/><a:ext cx="11277600" cy="731520"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="173B55"/></a:solidFill></p:spPr>\n'
                '      <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>'
                f'<a:p><a:r><a:rPr lang="en-US" sz="2200" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>'
                f"<a:t>{escape(title)}</a:t></a:r></a:p></p:txBody></p:sp>\n"
                '    <p:sp><p:nvSpPr><p:cNvPr id="3" name="PlacementWireframe"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>\n'
                '      <p:spPr><a:xfrm><a:off x="457200" y="1219200"/><a:ext cx="6400800" cy="4876800"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EDF2F7"/></a:solidFill>'
                '<a:ln w="19050"><a:solidFill><a:srgbClr val="173B55"/></a:solidFill></a:ln></p:spPr>\n'
                '      <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>'
                f'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1600" b="1"><a:solidFill><a:srgbClr val="173B55"/></a:solidFill></a:rPr>'
                f"<a:t>{escape(placement_banner)}</a:t></a:r></a:p></p:txBody></p:sp>\n"
                '    <p:sp><p:nvSpPr><p:cNvPr id="4" name="MetadataPanel"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>\n'
                '      <p:spPr><a:xfrm><a:off x="7162800" y="1219200"/><a:ext cx="4572000" cy="4876800"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
                '<a:ln w="12700"><a:solidFill><a:srgbClr val="CBD5E0"/></a:solidFill></a:ln></p:spPr>\n'
                f"      <p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody></p:sp>\n"
                "  </p:spTree></p:cSld>\n</p:sld>"
            )
            zf.writestr(f"ppt/slides/slide{i}.xml", slide_xml)
            zf.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slide_rels)
