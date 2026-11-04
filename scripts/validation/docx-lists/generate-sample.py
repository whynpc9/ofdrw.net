#!/usr/bin/env python3
"""Write a deterministic DOCX that exercises Native/BuiltIn list markers.

Usage: generate-sample.py OUTPUT.docx

Covers a 3-level legal list (%1. / %1.%2. / %1.%2.%3.) with mixed Chinese and
English text and local bold/italic runs, an upperRoman list, a chineseCounting
"%1、" list, a bullet list, one numId=0 paragraph, and enough items to reach a
second page. Standard library only; fixed ZIP timestamps.
"""
import sys
import zipfile
from xml.sax.saxutils import escape

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STAMP = (2026, 1, 1, 0, 0, 0)


def run(text, bold=False, italic=False):
    props = ("<w:b/>" if bold else "") + ("<w:i/>" if italic else "")
    props = f"<w:rPr>{props}</w:rPr>" if props else ""
    return f'<w:r>{props}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(runs, num_id=None, level=0, heading=False):
    ppr = ""
    if num_id is not None:
        ppr = f'<w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>'
    if heading:
        ppr += '<w:spacing w:after="120"/>'
    return f"<w:p><w:pPr>{ppr}</w:pPr>{''.join(runs)}</w:p>"


def body():
    out = [para([run("List numbering acceptance 列表编号验收", bold=True)], heading=True)]
    clause = 0
    for chapter in range(1, 4):
        out.append(para([run(f"Chapter {chapter} 第{chapter}章 general provisions")], 1, 0))
        for section in range(1, 4):
            runs = [run(f"Section {chapter}.{section} 条款内容 ")]
            if chapter == 1 and section == 2:
                runs.append(run("bold term 加粗", bold=True))
                runs.append(run(" then plain text."))
            elif chapter == 2 and section == 1:
                runs.append(run("italic term", italic=True))
                runs.append(run(" then plain 普通文本."))
            else:
                runs.append(run("ordinary wording."))
            out.append(para(runs, 1, 1))
            for item in range(1, 3):
                clause += 1
                out.append(para([run(f"Item {chapter}.{section}.{item} 细则 {clause}")], 1, 2))
    out.append(para([run("Roman list 罗马数字")], heading=True))
    for index in range(1, 6):
        out.append(para([run(f"Roman entry {index} 罗马条目")], 2, 0))
    out.append(para([run("Chinese counting list 中文编号")], heading=True))
    for index in range(1, 21):
        out.append(para([run(f"中文条目 Chinese entry {index}")], 3, 0))
    out.append(para([run("Bullet list 项目符号")], heading=True))
    for index in range(1, 5):
        out.append(para([run(f"Bullet entry {index} 符号条目")], 4, 0))
    out.append(para([run("numId=0 paragraph: no marker 无编号段落")], 0, 0))
    out.append('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>')
    return "".join(out)


def level(index, fmt, text, start=1):
    return (f'<w:lvl w:ilvl="{index}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{escape(text)}"/></w:lvl>')


NUMBERING = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="{W}">'
    '<w:abstractNum w:abstractNumId="1">' + level(0, "decimal", "%1.") + level(1, "decimal", "%1.%2.")
    + level(2, "decimal", "%1.%2.%3.") + "</w:abstractNum>"
    '<w:abstractNum w:abstractNumId="2">' + level(0, "upperRoman", "%1.") + "</w:abstractNum>"
    '<w:abstractNum w:abstractNumId="3">' + level(0, "chineseCounting", "%1、") + "</w:abstractNum>"
    '<w:abstractNum w:abstractNumId="4">' + level(0, "bullet", "•") + "</w:abstractNum>"
    + "".join(f'<w:num w:numId="{n}"><w:abstractNumId w:val="{n}"/></w:num>' for n in range(1, 5))
    + "</w:numbering>")

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    "</Types>")

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>")

DOCUMENT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    "</Relationships>")


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="{W}"><w:body>{body()}</w:body></w:document>')
    parts = [("[Content_Types].xml", CONTENT_TYPES), ("_rels/.rels", ROOT_RELS),
             ("word/_rels/document.xml.rels", DOCUMENT_RELS), ("word/document.xml", document),
             ("word/numbering.xml", NUMBERING)]
    with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in parts:
            info = zipfile.ZipInfo(name, date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, text.encode("utf-8"))


if __name__ == "__main__":
    main()
