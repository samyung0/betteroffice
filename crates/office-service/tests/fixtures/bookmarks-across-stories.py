"""Writes bookmarks-across-stories.docx: bookmarks in the body, a header and a
footnote, a REF field to the body's bookmark and a PAGE field in the header,
for the export golden test (tests/export_golden.rs).

usage: python bookmarks-across-stories.py  (writes the .docx next to this script)
"""

import os
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O = "application/vnd.openxmlformats-officedocument.wordprocessingml"


def field(code, result):
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve">{code}</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f"<w:r><w:t>{result}</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


PARTS = {
    "[Content_Types].xml": f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{O}.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="{O}.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="{O}.footnotes+xml"/></Types>',
    "_rels/.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>',
    "word/_rels/document.xml.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}/header" Target="header1.xml"/><Relationship Id="rId2" Type="{R}/footnotes" Target="footnotes.xml"/></Relationships>',
    "word/document.xml": f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>'
    '<w:p><w:bookmarkStart w:id="1" w:name="intro"/><w:r><w:t xml:space="preserve">Body text </w:t></w:r><w:bookmarkEnd w:id="1"/><w:r><w:footnoteReference w:id="1"/></w:r></w:p>'
    f'<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>{field(" REF intro " + chr(92) + "h ", "Body text")}</w:p>'
    '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
    "word/header1.xml": f'<w:hdr xmlns:w="{W}"><w:p><w:bookmarkStart w:id="2" w:name="head"/><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:bookmarkEnd w:id="2"/>{field(" PAGE ", "1")}</w:p></w:hdr>',
    "word/footnotes.xml": f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p><w:bookmarkStart w:id="3" w:name="note"/><w:r><w:t>A note</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p></w:footnote></w:footnotes>',
}

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bookmarks-across-stories.docx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as package:
    for name, xml in PARTS.items():
        package.writestr(zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0)), xml)
print(out)
