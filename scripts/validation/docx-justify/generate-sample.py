#!/usr/bin/env python3
"""Write a deterministic DOCX that exercises Native OFD justified paragraphs.

Usage: generate-sample.py OUTPUT.docx

Covers justified English, Chinese and mixed paragraphs with bold, italic and
underlined runs, centered and right-aligned contrast paragraphs, a justified
paragraph inside a table cell and enough text for two pages. Standard library
only; fixed ZIP timestamps.
"""
import sys
import zipfile
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STAMP = (2026, 1, 1, 0, 0, 0)

ENGLISH = ("Justified text spreads the remaining space of every wrapped line across its word gaps, so the "
           "right edge of the paragraph forms a straight column while the final line keeps its natural ragged end. ")
CHINESE = ("两端对齐是中文公文和商务文档中最常见的正文对齐方式，每一个自动换行的行都应当填满版心宽度，"
           "使右侧边缘整齐划一，而段落的最后一行保持左对齐，不做拉伸处理。")
MIXED = ("混合文本 mixed text 中英文混排时，OFD native 输出应当在空格处分配剩余空间，"
         "for example version 1.0 release notes 以及 PDF export 链路。")


def run(text, bold=False, italic=False, underline=False):
    props = ("<w:b/>" if bold else "") + ("<w:i/>" if italic else "") + ('<w:u w:val="single"/>' if underline else "")
    props = f"<w:rPr>{props}</w:rPr>" if props else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(alignment, *runs):
    ppr = f'<w:pPr><w:jc w:val="{alignment}"/><w:spacing w:after="120"/></w:pPr>' if alignment else ""
    return f"<w:p>{ppr}{''.join(runs)}</w:p>"


def table():
    cell_text = para("both", run("Cell paragraph 单元格内两端对齐 " * 6))
    other = para(None, run("Left cell 左侧单元格"))
    return ('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>'
            + "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
            + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>'
            + f'<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{other}</w:tc>'
            + f'<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{cell_text}</w:tc></w:tr></w:tbl>')


def body():
    out = [para(None, run("Justify acceptance 两端对齐验收", bold=True)),
           para("both", run(ENGLISH), run("Bold words ", bold=True), run("and "), run("italic words ", italic=True),
                run("and an "), run("underlined phrase that should stretch", underline=True), run(" " + ENGLISH)),
           para("both", run(CHINESE), run("加粗文字", bold=True), run(CHINESE)),
           para("both", run(MIXED), run("italic 斜体", italic=True), run(MIXED)),
           para("center", run("Centered contrast paragraph 居中对比段落")),
           para("right", run("Right-aligned contrast paragraph 右对齐对比段落")),
           table()]
    for index in range(1, 13):
        out.append(para("both", run(f"[{index}] "), run(ENGLISH if index % 2 else CHINESE * 2)))
    out.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>')
    return "".join(out)


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    "</Types>")

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>")


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="{W}"><w:body>{body()}</w:body></w:document>')
    parts = [("[Content_Types].xml", CONTENT_TYPES), ("_rels/.rels", ROOT_RELS), ("word/document.xml", document)]
    with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in parts:
            info = zipfile.ZipInfo(name, date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, text.encode("utf-8"))


if __name__ == "__main__":
    main()
