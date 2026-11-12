# -*- coding: utf-8 -*-
"""Rebuild the EN manuscript .docx from markdown, styled after the group's
reference manuscript docx.

Format spec implemented (see ../references/word-format-spec.md within this skill):
- A4, margins L/R 1800 twips (3.17 cm), T/B 1440 twips (2.54 cm)
- Body: Times New Roman 10.5 pt, double-spaced, justified, first-line indent
  210 twips; headings before=240; title/H1 12 pt bold, H2 10.5 pt bold
- Figures centered; captions (Figure/Table N.) 9 pt plain, no indent, and the
  leading "Figure N."/"Table N." label bold
- Inline math inside captions forced to the caption size (9 pt)
- Display equations: equation centered via center tab, number "(n)" flush
  right via right tab (extracted from the in-math \\qquad\\text{(n)} form)
- Tables full 0.5 pt grid
- First occurrence of each figure reference in body text bolded (done at the
  markdown preprocessing stage)
- H5+ headings (##### …) are author-only writing hints — stripped entirely
  (text plus line) at the markdown preprocessing stage, leaving no blank line

Usage: python _rebuild_word.py <in.md> <out.docx>   (script lives in word-standards skill scripts/; run from the manuscript dir)
"""
import os
import re
import subprocess
import sys
import zipfile

TNR = 'Times New Roman'
RF = '<w:rFonts w:ascii="{t}" w:hAnsi="{t}" w:eastAsia="{t}" w:cs="{t}"/>'.format(t=TNR)

DOCPR = ('<w:docDefaults><w:rPrDefault><w:rPr>' + RF +
         '<w:kern w:val="2"/><w:sz w:val="21"/><w:szCs w:val="22"/>'
         '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>'
         '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>')

SECTPR = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
          '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
          'w:header="851" w:footer="992" w:gutter="0"/>'
          '<w:cols w:space="425"/><w:docGrid w:linePitch="312"/></w:sectPr>')

GRID = ('<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        '</w:tblBorders>')

EQ_TABS = ('<w:tabs><w:tab w:val="center" w:pos="4153"/>'
           '<w:tab w:val="right" w:pos="8306"/></w:tabs>')
EQ_NUM_RPR = '<w:rPr>' + RF + '<w:sz w:val="21"/><w:szCs w:val="22"/></w:rPr>'
CAP_SZ = '<w:sz w:val="18"/><w:szCs w:val="18"/>'

# text width: A4 (11906 twips) - margins (1800+1800) = 8306 twips -> EMU
TEXT_W_EMU = 8306 * 635


# ---------------------------------------------------------------- styles ---
def style(sid, name, ppr='', rpr='', extra='', based=None, nxt=None,
          typ='paragraph'):
    x = ['<w:style w:type="' + typ + '"']
    if sid == 'Normal':
        x.append(' w:default="1"')
    x.append(' w:styleId="' + sid + '"><w:name w:val="' + name + '"/>')
    if based:
        x.append('<w:basedOn w:val="' + based + '"/>')
    if nxt:
        x.append('<w:next w:val="' + nxt + '"/>')
    x.append('<w:qFormat/>' + extra)
    if ppr:
        x.append('<w:pPr>' + ppr + '</w:pPr>')
    if rpr:
        x.append('<w:rPr>' + rpr + '</w:rPr>')
    x.append('</w:style>')
    return ''.join(x)


BODY_PPR = ('<w:spacing w:line="480" w:lineRule="auto"/>'
            '<w:ind w:firstLineChars="100" w:firstLine="210"/>'
            '<w:jc w:val="both"/>')
HEAD_RPR = RF + '<w:b/><w:bCs/>{sz}'

STYLES = {
    'Normal': style('Normal', 'Normal',
                    ppr='<w:widowControl w:val="0"/><w:jc w:val="both"/>'),
    'BodyText': style('BodyText', 'Body Text', ppr=BODY_PPR, based='Normal'),
    'FirstParagraph': style('FirstParagraph', 'First Paragraph',
                            ppr=BODY_PPR, based='Normal'),
    'Compact': style('Compact', 'Compact',
                     ppr='<w:spacing w:line="480" w:lineRule="auto"/>',
                     based='Normal'),
    'Heading1': style('Heading1', 'heading 1',
                      ppr=('<w:keepNext/><w:keepLines/>'
                           '<w:spacing w:before="240" w:after="120" '
                           'w:line="480" w:lineRule="auto"/>'
                           '<w:jc w:val="left"/><w:outlineLvl w:val="0"/>'),
                      rpr=HEAD_RPR.format(sz='<w:sz w:val="24"/><w:szCs w:val="24"/>'),
                      based='Normal', nxt='BodyText'),
    'Heading2': style('Heading2', 'heading 2',
                      ppr=('<w:keepNext/><w:keepLines/>'
                           '<w:spacing w:before="240" w:after="120" '
                           'w:line="480" w:lineRule="auto"/>'
                           '<w:jc w:val="left"/><w:outlineLvl w:val="1"/>'),
                      rpr=HEAD_RPR.format(sz='<w:sz w:val="24"/><w:szCs w:val="24"/>'),
                      based='Normal', nxt='BodyText'),
    'Heading3': style('Heading3', 'heading 3',
                      ppr=('<w:keepNext/><w:keepLines/>'
                           '<w:spacing w:before="240" w:after="120" '
                           'w:line="480" w:lineRule="auto"/>'
                           '<w:jc w:val="left"/><w:outlineLvl w:val="2"/>'),
                      rpr=HEAD_RPR.format(sz='<w:sz w:val="21"/><w:szCs w:val="22"/>'),
                      based='Normal', nxt='BodyText'),
    'ImageCaption': style('ImageCaption', 'Image Caption',
                          ppr=('<w:spacing w:line="360" w:lineRule="auto"/>'
                               '<w:ind w:firstLine="0" w:firstLineChars="0"/>'
                               '<w:jc w:val="both"/>'),
                          rpr=CAP_SZ, based='Normal'),
    'TableCaption': style('TableCaption', 'Table Caption',
                          ppr=('<w:keepNext/><w:spacing w:line="360" '
                               'w:lineRule="auto"/>'
                               '<w:ind w:firstLine="0" w:firstLineChars="0"/>'
                               '<w:jc w:val="both"/>'),
                          rpr=CAP_SZ, based='Normal'),
    'Table': style('Table', 'Table',
                   extra='<w:tblPr>' + GRID +
                         '<w:tblCellMar><w:top w:w="30" w:type="dxa"/>'
                         '<w:left w:w="80" w:type="dxa"/>'
                         '<w:bottom w:w="30" w:type="dxa"/>'
                         '<w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>',
                   typ='table'),
    # 参考文献条目样式（提取自参考稿 styles.xml 的 EndNoteBibliography：
    # TNR 10pt、双倍行距、无首行缩进、两端对齐）
    'EndNoteBibliography': style(
        'EndNoteBibliography', 'EndNote Bibliography',
        ppr=('<w:spacing w:line="480" w:lineRule="auto"/>'
             '<w:ind w:firstLine="0" w:firstLineChars="0"/>'
             '<w:jc w:val="both"/>'),
        rpr=RF + '<w:sz w:val="20"/><w:szCs w:val="20"/>',
        based='Normal', nxt='EndNoteBibliography'),
}


def patch_styles(styles: str) -> str:
    styles = re.sub(r'<w:docDefaults>.*?</w:docDefaults>', DOCPR, styles,
                    flags=re.S)
    for sid, xml in STYLES.items():
        pat = r'<w:style [^>]*w:styleId="' + sid + r'"[^>]*>.*?</w:style>'
        if re.search(pat, styles, flags=re.S):
            styles = re.sub(pat, xml, styles, flags=re.S)
        else:
            styles = styles.replace('</w:styles>', xml + '</w:styles>')
    return styles


# --------------------------------------------------------- md preprocessing --
def bold_first_fig_refs(txt: str):
    """Wrap the first body-text reference to each figure in **bold**.
    Skips captions, images, tables, headings, display math, bold spans.
    Returns (new_txt, number_of_bolded_tokens)."""
    seen = set()
    nbold = [0]
    pat = re.compile(r'\b(Figure\s+\d+|Figs?\.\s*\d+)(\s*[–—-]\s*(\d+))?')

    def mark_nums(m):
        a = int(re.search(r'\d+', m.group(1)).group())
        b = int(m.group(3)) if m.group(3) else a
        return set(range(a, b + 1))

    def repl(m):
        nums = mark_nums(m)
        if nums & seen:
            return m.group(0)
        seen.update(nums)
        nbold[0] += 1
        return '**' + m.group(0) + '**'

    out, in_math = [], False
    for ln in txt.split('\n'):
        s = ln.lstrip()
        if s.startswith('$$'):
            in_math = not in_math
        if (in_math or s.startswith(('![', '|', '#')) or
                re.match(r'^\*{1,3}\s*(Figure|Table)\b', s) or
                re.match(r'^\*\*', s) or
                re.match(r'^(Figure|Table)\s+\d', s)):
            out.append(ln)
            continue
        out.append(pat.sub(repl, ln))
    return '\n'.join(out), nbold[0]


# ------------------------------------------------------------- xml helpers --
def refs_fence_to_paras(txt: str):
    """把 ```reference 围栏块转成 '## References' 标题 + 普通段落，
    避免 pandoc 将文献列表渲染成等宽代码块（参考稿为正文段落）。
    条目间插空行——否则 Markdown 相邻行会合并成一个段落。"""
    m = re.search(r'```reference\n(.*?)\n```', txt, flags=re.S)
    if not m:
        return txt, 0
    inner = m.group(1).strip()
    n = inner.count('\n') + 1
    inner = inner.replace('\n', '\n\n')
    return txt.replace(m.group(0), '## References\n\n' + inner), n


def style_references(doc: str):
    """References 标题之后的 [N] 条目段套 EndNoteBibliography 样式
    （10pt/双倍行距/无首行缩进，提取自参考稿）。返回 (doc, 套用条数)。"""
    paras = re.findall(r'<w:p\b.*?</w:p>', doc, flags=re.S)
    idx = None
    for i, p in enumerate(paras):
        if _plain(p) == 'References':
            idx = i
    if idx is None:
        return doc, 0
    n = 0
    for p in paras[idx + 1:]:
        if not re.match(r'^\s*\[', _plain(p)):
            continue
        np = p
        if '<w:pStyle ' in np:
            np = re.sub(r'<w:pStyle w:val="[^"]*"\s*/>',
                        '<w:pStyle w:val="EndNoteBibliography"/>', np,
                        count=1)
        else:
            np = np.replace('<w:pPr>',
                            '<w:pPr><w:pStyle w:val="EndNoteBibliography"/>',
                            1)
        doc = doc.replace(p, np, 1)
        n += 1
    return doc, n


def _plain(x: str) -> str:
    return re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', '', x)).strip()


def set_jc(p: str, val: str) -> str:
    pr = re.search(r'<w:pPr>(.*?)</w:pPr>', p, flags=re.S)
    if pr:
        inner = re.sub(r'<w:jc w:val="[^"]*"/>', '', pr.group(1))
        if '<w:rPr>' in inner:
            inner = inner.replace('<w:rPr>',
                                  '<w:jc w:val="' + val + '"/><w:rPr>', 1)
        else:
            inner += '<w:jc w:val="' + val + '"/>'
        return p[:pr.start(1)] + inner + p[pr.end(1):]
    return re.sub(r'(<w:p(?: [^>]*)?>)',
                  lambda m: m.group(1) +
                  '<w:pPr><w:jc w:val="' + val + '"/></w:pPr>', p, count=1)


def bold_caption_label(q: str) -> str:
    """Bold the leading 'Figure N.'/'Table N.' label of a caption paragraph."""
    mm = re.search(r'(<w:r>)<w:rPr>((?:(?!</w:rPr>).)*)</w:rPr>'
                   r'<w:t[^>]*>([^<]*)</w:t></w:r>', q, flags=re.S)
    if not mm:
        return q
    lm = re.match(r'^((?:Figure|Table)\s+\d+)(\.?)\s*(.*)$', mm.group(3))
    if not lm:
        return q
    rpr = mm.group(2)
    runs = ['<w:r><w:rPr><w:b/><w:bCs/>' + rpr + '</w:rPr>'
            '<w:t xml:space="preserve">' + lm.group(1) + lm.group(2) +
            '</w:t></w:r>']
    if lm.group(3):
        runs.append('<w:r><w:rPr>' + rpr + '</w:rPr>'
                    '<w:t xml:space="preserve">' + lm.group(3) +
                    '</w:t></w:r>')
    return q[:mm.start()] + ''.join(runs) + q[mm.end():]


def size_caption_math(q: str) -> str:
    """Give every math run in a caption the caption size (9 pt). OMML m:r uses
    m:rPr for math styling; the word-level w:rPr (which controls size) may be
    absent -> insert it right after <m:r> (schema: w:rPr precedes m:rPr)."""
    q = re.sub(r'(<m:r>)(?!<w:rPr>|<m:rPr>)',
               r'\1<w:rPr>' + CAP_SZ + '</w:rPr>', q)
    q = re.sub(r'(<m:r>)(<m:rPr>)',
               r'\1<w:rPr>' + CAP_SZ + r'</w:rPr>\2', q)
    return q


def captionize(p: str) -> str:
    q = re.sub(r'<w:(b|bCs|i|iCs)\s*/>', '', p)
    q = re.sub(r'<w:sz w:val="\d+"\s*/>', '', q)
    q = re.sub(r'<w:szCs w:val="\d+"\s*/>', '', q)

    def add_sz(mm):
        inner = mm.group(1)
        if '<w:lang' in inner:
            return '<w:rPr>' + inner.replace(
                '<w:lang', CAP_SZ + '<w:lang', 1) + '</w:rPr>'
        return '<w:rPr>' + inner + CAP_SZ + '</w:rPr>'
    q = re.sub(r'<w:rPr>(.*?)</w:rPr>', add_sz, q, flags=re.S)
    q = re.sub(r'(<w:r>)(?!<w:rPr>)',
               r'\1<w:rPr>' + CAP_SZ + '</w:rPr>', q)
    q = set_jc(q, 'both')
    pr = re.search(r'<w:pPr>(.*?)</w:pPr>', q, flags=re.S)
    if pr:
        inner = re.sub(r'<w:ind [^>]*/>', '', pr.group(1))
        inner += '<w:ind w:firstLine="0" w:firstLineChars="0"/>'
        q = q[:pr.start(1)] + inner + q[pr.end(1):]
    q = size_caption_math(q)
    q = bold_caption_label(q)
    return q


def captions_sub(doc: str):
    """Style Figure/Table captions: 9 pt plain, no indent, bold label. A
    caption is a paragraph starting with 'Figure/Table N' immediately
    preceded by an image paragraph or followed by a table."""
    blocks = [(m.start(), m.end(), m.group(0))
              for m in re.finditer(
                  r'<w:p(?: [^>]*)?>.*?</w:p>|<w:tbl>.*?</w:tbl>', doc,
                  flags=re.S)]
    styled = []
    for i, (s, e, x) in enumerate(blocks):
        if not x.startswith('<w:p'):
            continue
        if not re.match(r'^(Figure|Table)\s+\d+', _plain(x)):
            continue
        is_fig = False
        j = i - 1
        while j >= 0:
            y = blocks[j][2]
            if y.startswith('<w:tbl'):
                break
            if '<w:drawing' in y:
                is_fig = not _plain(y)
                break
            if _plain(y):
                break
            j -= 1
        is_tbl = False
        k = i + 1
        while k < len(blocks):
            y = blocks[k][2]
            if y.startswith('<w:tbl'):
                is_tbl = True
                break
            if _plain(y):
                break
            k += 1
        if is_fig or is_tbl:
            styled.append(((s, e), captionize(x)))
    for (s, e), new in reversed(styled):
        doc = doc[:s] + new + doc[e:]
    return doc, len(styled)


def equations_sub(doc: str):
    """Display equations: extract the trailing in-math '(n)' number, strip the
    oMathPara wrapper (-> bare inline oMath, as in the reference manuscript),
    then rebuild the paragraph as  [tab->center] equation [tab->right] (n)."""
    stats = {'eq': 0, 'num': 0}

    def repl(m):
        p = m.group(0)
        if '<m:oMathPara' not in p:
            return p
        stats['eq'] += 1
        num = None
        for pat in (r'<m:r><m:t[^>]*>[^<]*</m:t></m:r>'
                    r'<m:r><m:rPr><m:nor\s*/><m:sty m:val="p"\s*/></m:rPr>'
                    r'<m:t[^>]*>\((\d+)\)</m:t></m:r></m:oMath>',
                    r'<m:r><m:rPr><m:nor\s*/><m:sty m:val="p"\s*/></m:rPr>'
                    r'<m:t[^>]*>\((\d+)\)</m:t></m:r></m:oMath>'):
            mm = re.search(pat, p)
            if mm:
                num = mm.group(1)
                p = p[:mm.start()] + '</m:oMath>' + p[mm.end():]
                stats['num'] += 1
                break
        p = re.sub(r'<m:oMathPara>(?:<m:oMathParaPr>.*?</m:oMathParaPr>)?'
                   r'(.*?)</m:oMathPara>', r'\1', p, flags=re.S)
        p = re.sub(r'<w:pPr>.*?</w:pPr>',
                   '<w:pPr><w:pStyle w:val="BodyText"/>' + EQ_TABS +
                   '<w:jc w:val="left"/></w:pPr>', p, count=1, flags=re.S)
        p = re.sub(r'(</w:pPr>)', r'\1<w:r><w:tab/></w:r>', p, count=1)
        if num:
            p = p.replace(
                '</m:oMath>',
                '</m:oMath><w:r>' + EQ_NUM_RPR + '<w:tab/>'
                '<w:t xml:space="preserve">(' + num + ')</w:t></w:r>', 1)
        return p
    doc = re.sub(r'<w:p(?: [^>]*)?>.*?</w:p>', repl, doc, flags=re.S)
    return doc, stats


def center_images(doc: str) -> str:
    def repl(m):
        p = m.group(0)
        if '<w:drawing' not in p:
            return p
        return set_jc(p, 'center')
    return re.sub(r'<w:p(?: [^>]*)?>.*?</w:p>', repl, doc, flags=re.S)


def grid_tables(doc: str) -> str:
    def repl(m):
        pr = re.sub(r'<w:tblBorders>.*?</w:tblBorders>', '', m.group(0),
                    flags=re.S)
        return pr.replace('</w:tblPr>', GRID + '</w:tblPr>')
    return re.sub(r'<w:tblPr>.*?</w:tblPr>', repl, doc, flags=re.S)


# --------------------------------------------------- dpi-aware image sizing --
def _img_dpi(data: bytes, name: str):
    """Return (w_px, h_px, dpi) from PNG pHYs / JPEG JFIF density metadata.
    Returns dpi=None when the file carries no usable physical-size info."""
    import struct

    if name.lower().endswith('.png'):
        w = h = dpi = None
        pos = 8
        while pos < len(data) - 8:
            ln = struct.unpack('>I', data[pos:pos + 4])[0]
            typ = data[pos + 4:pos + 8]
            body = data[pos + 8:pos + 8 + ln]
            if typ == b'IHDR':
                w, h = struct.unpack('>II', body[:8])
            if typ == b'pHYs' and len(body) >= 9:
                x = struct.unpack('>I', body[:4])[0]
                unit = body[8]
                if unit == 1 and x:
                    dpi = x * 0.0254          # px/m -> dpi
            pos += 12 + ln
        return w, h, dpi

    # jpeg: scan JFIF APP0 density + SOFn dims
    w = h = dpi = None
    pos = 2
    while pos < len(data) - 1:
        if data[pos] != 0xFF:
            pos += 1
            continue
        marker = data[pos + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            pos += 2
            continue
        ln = struct.unpack('>H', data[pos + 2:pos + 4])[0]
        seg = data[pos + 4:pos + 2 + ln]
        if (marker == 0xE0 and seg[:5] == b'JFIF\x00'
                and len(seg) >= 10):
            units = seg[7]
            xd = struct.unpack('>H', seg[8:10])[0]
            if xd and units == 1:
                dpi = float(xd)              # dots/inch
            elif xd and units == 2:
                dpi = xd * 2.54              # dots/cm -> dpi
        if marker in (0xC0, 0xC1, 0xC2, 0xC3):
            h = struct.unpack('>H', seg[1:3])[0]
            w = struct.unpack('>H', seg[3:5])[0]
        if marker == 0xDA:
            break
        pos += 2 + ln
    return w, h, dpi


def size_images(doc: str, media_by_rid: dict):
    """Set each inline image to its DPI-intrinsic physical size, capped at
    the text width. Images without DPI metadata keep pandoc's own sizing
    (96 dpi assumption + page-width cap)."""
    stats = {'dpi': 0, 'capped': 0, 'default': 0}
    sizes = []

    def repl(m):
        block = m.group(0)
        rid = re.search(r'r:embed="(rId\d+)"', block)
        if not rid:
            return block
        name, data = media_by_rid.get(rid.group(1), (None, None))
        if not data:
            return block
        descr = re.search(r'pic:cNvPr descr="([^"]+)"', block)
        label = (descr.group(1).split('/')[-1] if descr and descr.group(1)
                 else name or '?')
        w, h, dpi = _img_dpi(data, name or '')
        if not (w and h and dpi):
            stats['default'] += 1
            cx = re.search(r'<wp:extent cx="(\d+)"', block)
            sizes.append((label, round(int(cx.group(1)) / 360000, 1)
                          if cx else '?', 'no-dpi (pandoc default)'))
            return block
        cx = w / dpi * 914400.0
        cy = h / dpi * 914400.0
        if cx > TEXT_W_EMU:
            cy *= TEXT_W_EMU / cx
            cx = float(TEXT_W_EMU)
            stats['capped'] += 1
            note = 'dpi, capped to text width'
        else:
            stats['dpi'] += 1
            note = 'dpi'
        sizes.append((label, round(cx / 360000, 1), note))
        block = re.sub(r'(<wp:extent cx=")\d+(" cy=")\d+("\s*/>)',
                       lambda mm: mm.group(1) + str(int(cx)) + mm.group(2)
                       + str(int(cy)) + mm.group(3), block, count=1)
        block = re.sub(r'(<a:ext cx=")\d+(" cy=")\d+("\s*/>)',
                       lambda mm: mm.group(1) + str(int(cx)) + mm.group(2)
                       + str(int(cy)) + mm.group(3), block, count=1)
        return block

    doc = re.sub(r'<w:drawing>.*?</w:drawing>', repl, doc, flags=re.S)
    return doc, stats, sizes


# ------------------------------------------------------------------- main ---
def main(md_path, out_path):
    txt = open(md_path, encoding='utf-8').read()
    txt = re.sub(r'\s*\\tag\{([^}]+)\}', r'\\qquad\\text{(\1)}', txt)
    # ##### 及更深的标题仅是作者写作提示：整行剥离（含换行），不显示也不留空行
    txt, n_hint = re.subn(r'^#{5,}[^\n]*\n?', '', txt, flags=re.M)
    txt, n_refparas = refs_fence_to_paras(txt)
    txt, n_refs = bold_first_fig_refs(txt)
    tmp_md = '_tmp_build.md'
    open(tmp_md, 'w', encoding='utf-8').write(txt)
    print('reference fence -> paragraphs:', n_refparas)
    print('author-hint headings (#####) stripped:', n_hint)
    print('first-occurrence figure refs bolded in body:', n_refs)

    ref = subprocess.run(
        ['pandoc', '--print-default-data-file', 'reference.docx'],
        capture_output=True, check=True).stdout
    tmp_ref = '_tmp_ref.docx'
    open(tmp_ref, 'wb').write(ref)
    with zipfile.ZipFile(tmp_ref) as zin:
        entries = {i.filename: zin.read(i.filename) for i in zin.infolist()}
        infos = zin.infolist()
    entries['word/styles.xml'] = patch_styles(
        entries['word/styles.xml'].decode('utf-8')).encode('utf-8')
    d = entries['word/document.xml'].decode('utf-8')
    d = re.sub(r'<w:sectPr.*?</w:sectPr>', SECTPR, d, flags=re.S)
    entries['word/document.xml'] = d.encode('utf-8')
    with zipfile.ZipFile(tmp_ref, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info.filename, entries[info.filename])

    tmp_out = '_tmp_out.docx'
    subprocess.run(['pandoc', '-f', 'markdown-implicit_figures',
                    tmp_md, '-o', tmp_out,
                    '--reference-doc=' + tmp_ref,
                    '--resource-path=.', '--metadata', 'lang=en'],
                   check=True)

    with zipfile.ZipFile(tmp_out) as zin:
        entries = {i.filename: zin.read(i.filename) for i in zin.infolist()}
        infos = zin.infolist()
    doc = entries['word/document.xml'].decode('utf-8')
    n_math = doc.count('Cambria Math')
    doc, n_caps = captions_sub(doc)
    doc = center_images(doc)
    rels = entries['word/_rels/document.xml.rels'].decode('utf-8')
    media_by_rid = {}
    for rid, tgt in re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels):
        if 'media/' in tgt:
            key = 'word/' + tgt.lstrip('/')
            if key in entries:
                media_by_rid[rid] = (tgt.split('/')[-1], entries[key])
    doc, img_stats, img_sizes = size_images(doc, media_by_rid)
    doc, eq = equations_sub(doc)
    doc = grid_tables(doc)
    doc, n_refstyled = style_references(doc)
    print('reference paragraphs styled:', n_refstyled)
    entries['word/document.xml'] = doc.encode('utf-8')
    with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info.filename, entries[info.filename])

    # ---- verify ----
    import xml.dom.minidom
    with zipfile.ZipFile(tmp_out) as z:
        d = z.read('word/document.xml').decode('utf-8')
        s = z.read('word/styles.xml').decode('utf-8')
    xml.dom.minidom.parseString(d.encode('utf-8'))
    xml.dom.minidom.parseString(s.encode('utf-8'))
    print('XML well-formed: True')
    paras = re.findall(r'<w:p(?: [^>]*)?>.*?</w:p>', d, flags=re.S)
    imgs = [p for p in paras if '<w:drawing' in p]
    imgs_c = sum('<w:jc w:val="center"/>' in p for p in imgs)
    tbls = len(re.findall(r'<w:tbl>', d))
    tbls_b = len(re.findall(r'<w:tblBorders>', d))
    caps = [p for p in paras
            if re.match(r'^(Figure|Table)\s+\d+', _plain(p))
            and CAP_SZ in p]
    caps_b = sum(bool(re.search(r'<w:r><w:rPr><w:b/><w:bCs/>', p))
                 for p in caps)
    capmath_ok = all(
        len(re.findall(r'<m:r>', p)) ==
        len(re.findall(r'<m:r><w:rPr>' + CAP_SZ, p)) for p in caps)
    eq_paras = [p for p in paras if '<w:tab/></w:r><m:oMath' in p]
    eq_num = sum(bool(re.search(
        r'<w:tab/><w:t xml:space="preserve">\(\d+\)</w:t></w:r>', p))
        for p in eq_paras)
    omathpara_left = d.count('<m:oMathPara')
    math_ok = d.count('Cambria Math') == n_math
    marg_ok = 'w:left="1800"' in d and 'w:top="1440"' in d
    bold_refs = sum(bool(re.search(
        r'<w:r><w:rPr>(?:(?!</w:rPr>).)*<w:b\s*/>(?:(?!</w:rPr>).)*</w:rPr>'
        r'<w:t[^>]*>\s*(Figure|Fig\.|Figs\.)', p, flags=re.S))
        for p in paras if not re.match(r'^(Figure|Table)\s+\d+', _plain(p)))
    print('images: {} total, {} centered'.format(len(imgs), imgs_c))
    print('image sizing: {} by dpi, {} dpi-capped, {} no-dpi default'
          .format(img_stats['dpi'], img_stats['capped'],
                  img_stats['default']))
    for name, cm, note in img_sizes:
        print('  {:<32} {:>6} cm  ({})'.format(name, cm, note))
    print('tables: {} total, {} bordered'.format(tbls, tbls_b))
    print('captions: {} styled, {} bold-label, math-sized ok: {}'
          .format(len(caps), caps_b, capmath_ok))
    print('display eqs: {} total, {} numbered-right, oMathPara left: {}'
          .format(len(eq_paras), eq_num, omathpara_left))
    print('body bold fig-ref paragraphs:', bold_refs,
          '(expected {} bolded tokens)'.format(n_refs))
    print('Cambria Math preserved:', math_ok, '| margins ok:', marg_ok)
    ok = (imgs and imgs_c == len(imgs) and tbls_b == tbls
          and len(caps) == 17 and caps_b == 17 and capmath_ok
          and len(eq_paras) == 20 and eq_num == 17 and omathpara_left == 0
          and bold_refs >= n_refs and math_ok and marg_ok)
    if not ok:
        print('VERIFICATION FAILED; inspect _tmp_out.docx')
        sys.exit(1)

    try:
        os.replace(tmp_out, out_path)
        print('OK, rebuilt:', out_path)
    except PermissionError:
        alt = out_path.replace('.docx', '-new.docx')
        os.replace(tmp_out, alt)
        print('TARGET LOCKED (close Word?). Wrote instead:', alt)
    finally:
        for f in (tmp_md, tmp_ref):
            if os.path.exists(f):
                os.remove(f)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
