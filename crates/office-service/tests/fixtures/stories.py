"""Writes stories.docx: a header, a footer, a footnote, a table with a nested
table, a block date control (Japanese), a block rich-text control and a page
break, for the DOCX story walk tests (src/docx/reach.rs).

usage: python stories.py  (writes stories.docx next to this script)
"""

import os
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def p(text, extra=""):
    return f'<w:p>{extra}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def cell(content):
    return f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>{content}</w:tc>'


nested = f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/></w:tblGrid><w:tr>{cell(p("Nested"))}</w:tr></w:tbl>{p("")}'
table = f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>{cell(p("Cell"))}{cell(nested)}</w:tr></w:tbl>'
date = (
    '<w:sdt><w:sdtPr><w:id w:val="11"/><w:date w:fullDate="2026-03-04T00:00:00Z">'
    '<w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="ja-JP"/><w:storeMappedDataAs w:val="dateTime"/>'
    '<w:calendar w:val="gregorian"/></w:date></w:sdtPr><w:sdtContent>'
    + p("4 3月 2026")
    + "</w:sdtContent></w:sdt>"
)
rich = '<w:sdt><w:sdtPr><w:id w:val="12"/></w:sdtPr><w:sdtContent>' + p("Rich text") + "</w:sdtContent></w:sdt>"
footnote_ref = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'
body = (
    f'<w:p><w:r><w:t xml:space="preserve">Intro with a note</w:t></w:r>{footnote_ref}</w:p>'
    + table
    + date
    + rich
    + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
    + p("End")
    + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>'
    '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)
parts = {
    "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>',
    "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    "word/_rels/document.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>',
    "word/document.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>',
    "word/header1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{p("Header")}</w:hdr>',
    "word/footer1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}>{p("Footer")}</w:ftr>',
    "word/footnotes.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">{p("A note")}</w:footnote></w:footnotes>',
}
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "stories.docx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as package:
    for name, xml in parts.items():
        info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        package.writestr(info, xml.encode("utf-8"))
print(out)
