# -*- coding: utf-8 -*-
"""Patch a pandoc/Word .docx: Times New Roman everywhere, centered images,
gridded tables. Usage: python _fix_docx_format.py <file.docx>

- Fonts: replaces every explicit <w:rFonts> with Times New Roman and keeps
  math runs (Cambria Math inside <m:oMath>) untouched so equations still render.
- Images: every paragraph containing <w:drawing> gets w:jc="center".
- Tables: every table gets explicit single-line borders (all six edges).
Operates on a temp copy, verifies, then atomically replaces the original.
"""
import os
import re
import sys
import zipfile

TNR = "Times New Roman"
RFONT = ('<w:rFonts w:ascii="{t}" w:hAnsi="{t}" w:eastAsia="{t}" w:cs="{t}"/>'
         .format(t=TNR))

BORDERS = (
    '<w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '</w:tblBorders>'
)


def set_fonts_outside_math(doc: str) -> str:
    """Replace rFonts with TNR everywhere except inside <m:oMath> math zones."""
    parts = re.split(r'(<m:oMath(?: [^>]*)?>.*?</m:oMath>)', doc, flags=re.S)
    for i in range(0, len(parts), 2):          # even indexes: outside math
        parts[i] = re.sub(r'<w:rFonts[^>]*/>', RFONT, parts[i])
    return ''.join(parts)


def center_image_paragraphs(doc: str) -> str:
    def fix(m):
        p = m.group(0)
        if '<w:drawing' not in p:
            return p
        pr = re.search(r'<w:pPr>(.*?)</w:pPr>', p, flags=re.S)
        if pr:
            inner = re.sub(r'<w:jc w:val="[^"]*"/>', '', pr.group(1))
            if '<w:rPr>' in inner:              # jc must precede rPr (schema order)
                inner = inner.replace('<w:rPr>', '<w:jc w:val="center"/><w:rPr>', 1)
            else:
                inner += '<w:jc w:val="center"/>'
            return p[:pr.start(1)] + inner + p[pr.end(1):]
        return re.sub(r'(<w:p(?: [^>]*)?>)',
                      lambda mm: mm.group(1) + '<w:pPr><w:jc w:val="center"/></w:pPr>',
                      p, count=1)
    return re.sub(r'<w:p(?: [^>]*)?>.*?</w:p>', fix, doc, flags=re.S)


def grid_tables(doc: str) -> str:
    def fix(m):
        pr = m.group(0)
        pr = re.sub(r'<w:tblBorders>.*?</w:tblBorders>', '', pr, flags=re.S)
        return pr.replace('</w:tblPr>', BORDERS + '</w:tblPr>')
    return re.sub(r'<w:tblPr>.*?</w:tblPr>', fix, doc, flags=re.S)


def main(path: str) -> None:
    tmp = path + '.tmpfmt'
    with zipfile.ZipFile(path) as zin:
        entries = {i.filename: zin.read(i.filename) for i in zin.infolist()}
        infos = zin.infolist()

    styles = entries['word/styles.xml'].decode('utf-8')
    doc = entries['word/document.xml'].decode('utf-8')

    n_math = doc.count('Cambria Math')
    styles = re.sub(r'<w:rFonts[^>]*/>', RFONT, styles)
    doc = set_fonts_outside_math(doc)
    doc = center_image_paragraphs(doc)
    doc = grid_tables(doc)

    entries['word/styles.xml'] = styles.encode('utf-8')
    entries['word/document.xml'] = doc.encode('utf-8')

    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info.filename, entries[info.filename])

    # verify the patched copy before replacing the original
    with zipfile.ZipFile(tmp) as z:
        d = z.read('word/document.xml').decode('utf-8')
        s = z.read('word/styles.xml').decode('utf-8')
    imgs = re.findall(r'<w:p(?: [^>]*)?>.*?</w:p>', d, flags=re.S)
    imgs = [p for p in imgs if '<w:drawing' in p]
    centered = sum('<w:jc w:val="center"/>' in p for p in imgs)
    tbl_total = len(re.findall(r'<w:tbl>', d))
    tbl_bordered = len(re.findall(r'<w:tblBorders>', d))
    math_ok = d.count('Cambria Math') == n_math
    print('images: {} total, {} centered'.format(len(imgs), centered))
    print('tables: {} total, {} bordered'.format(tbl_total, tbl_bordered))
    print('Cambria Math runs preserved:', math_ok)
    print('TNR font tags in doc/styles:', d.count(TNR) + s.count(TNR))
    if not (imgs and centered == len(imgs) and tbl_bordered == tbl_total and math_ok):
        print('VERIFICATION FAILED -- original left untouched:', tmp)
        sys.exit(1)

    os.replace(tmp, path)
    print('OK, patched:', path)


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit('usage: python _fix_docx_format.py <file.docx>')
    main(sys.argv[1])
