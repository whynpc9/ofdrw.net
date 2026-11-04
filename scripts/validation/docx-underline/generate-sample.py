#!/usr/bin/env python3
"""Write a deterministic DOCX that exercises Native OFD underlines.

Usage: generate-sample.py OUTPUT.docx

Covers mid-line underlined words in mixed Chinese/English text, an underlined
bold run, an underlined red run, a run that wraps across lines, a fill-in blank
of underlined spaces, a 2x2 table with an underlined cell word and a footer with
underlined text, over two pages. Standard library only; fixed ZIP timestamps.
"""
import sys
import zipfile
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
STAMP = (2026, 1, 1, 0, 0, 0)


def run(text, underline=False, bold=False, color=None):
    props = ("<w:b/>" if bold else "") + (f'<w:color w:val="{color}"/>' if color else "")
    props += '<w:u w:val="single"/>' if underline else ""
    props = f"<w:rPr>{props}</w:rPr>" if props else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(*runs):
    return f"<w:p>{''.join(runs)}</w:p>"


def cell(*runs):
    return ('<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>'
            f"{para(*runs)}</w:tc>")


TABLE = (
    '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>'
    + "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
    + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>'
    + "<w:tr>" + cell(run("项目 Item")) + cell(run("说明 "), run("underlined", underline=True), run(" cell")) + "</w:tr>"
    + "<w:tr>" + cell(run("金额 Amount")) + cell(run("1,234.50 元")) + "</w:tr></w:tbl>")


def body():
    out = [para(run("Underline acceptance 下划线验收", bold=True)),
           para(run("This sentence has an "), run("underlined", underline=True),
                run(" word in the middle and 中文"), run("下划线", underline=True), run("文字 continues.")),
           para(run("Bold underline: "), run("加粗下划线 Bold", underline=True, bold=True),
                run(" and red underline: "), run("红色 Red", underline=True, color="C00000"), run(" end.")),
           para(run("Wrapping: "), run("this underlined run is deliberately long so that it wraps across the "
                                       "right margin onto the next line 并且包含中文字符以检查换行后的下划线", underline=True),
                run(" after.")),
           para(run("姓名："), run("                    ", underline=True), run("  日期："),
                run("            ", underline=True)),
           TABLE]
    for index in range(1, 46):
        if index % 9 == 0:
            out.append(para(run(f"Filler {index} 填充段落 with "), run(f"link {index}", underline=True), run(" inside.")))
        else:
            out.append(para(run(f"Filler paragraph {index} 填充段落内容用于分页 ordinary text without underline.")))
    out.append('<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
               '<w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>'
               "</w:sectPr>")
    return "".join(out)


FOOTER = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="{W}">'
          + para(run("Footer 页脚 "), run("underlined footer", underline=True), run(" text")) + "</w:ftr>")

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    "</Types>")

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>")

DOCUMENT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    "</Relationships>")


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body()}</w:body></w:document>')
    parts = [("[Content_Types].xml", CONTENT_TYPES), ("_rels/.rels", ROOT_RELS),
             ("word/_rels/document.xml.rels", DOCUMENT_RELS), ("word/document.xml", document),
             ("word/footer1.xml", FOOTER)]
    with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in parts:
            info = zipfile.ZipInfo(name, date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, text.encode("utf-8"))


if __name__ == "__main__":
    main()
