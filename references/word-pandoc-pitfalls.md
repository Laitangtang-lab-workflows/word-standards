# Word generation pitfalls (the full collection)

**Scenario**: a manuscript written in Chinese md, translated into English md, then built into Word styled after the research group's reference manuscript.
**Toolchain**: pandoc 3.x + Python 3.12 (zipfile reading and writing the docx XML directly).
**Reusable scripts** (shipped in this skill's `scripts\`): `_rebuild_word.py` (full md→docx pipeline), `_fix_docx_format.py` (patches three formatting items in place on a hand-edited docx).

## 1. Pandoc-layer pitfalls

### 1.1 `\tag{n}` silently dropped
pandoc/texmath **silently discards** `\tag{1}` when converting to OMML — no error, no trace, the equation renders fine, only the number is gone.
**Countermeasure**: before generating, replace `\s*\\tag\{([^}]+)\}` with `\\qquad\\text{(\1)}` so the number rides inside the equation's right end. The md source keeps the standard `\tag` spelling untouched.

### 1.2 Implicit figures: `![non-empty alt](img)` generates duplicate captions
For an image in its own paragraph with non-empty alt, pandoc auto-generates a figure **with a caption built from the alt text**, duplicating the hand-written caption.
**Countermeasure**: convert with `-f markdown-implicit_figures`; leave image alt empty in md.

### 1.3 `$1` vs `\1`: Python and PowerShell backreferences differ ⚠️ the biggest trap
Python `re.sub` backreferences are `\1`/`\g<1>`; PowerShell `-replace` and JS use `$1`. Carrying the PowerShell habit into Python outputs the **literal `($1)`** — 17 equations all failed to parse (caught only because pandoc happened to warn `unexpected "}"`).
**Countermeasure**: when writing regex across languages, run a minimal example first and check the replacement output.

### 1.4 The reference.docx mechanism
pandoc inherits **style definitions and sectPr (page size and margins)** from `--reference-doc`, but the body content structure is still decided by the writer. The correct play: `pandoc --print-default-data-file reference.docx` to get the default template → edit its styles.xml (docDefaults + the individual style blocks) and document.xml's sectPr with a script → convert using it. Editing the body XML is worse than editing the style sheet.

### 1.5 Chinese filenames in image paths are fine
`![](./figures/黏性主导状态仿真形貌相图.jpg)` works directly, provided pandoc runs **in the manuscript directory** (relative paths + `--resource-path=.`). Note the `./figures/` prefix; jpg or png on disk both work.

## 2. OOXML / regex-layer pitfalls

### 2.1 Attribute groups must be optional: `<w:p(?: [^>]*)?>` not `<w:p(?: [^>]*)>`
OOXML tags may or may not carry attributes (`<m:oMath>` vs `<m:rPr w:rsidR=...>`). Omitting the `?` on `(?: [^>]*)` turns it into "at least one attribute required" — zero matches on 586 bare `<m:oMath>` with no error whatsoever. This once nearly replaced every Cambria Math with TNR.

### 2.2 Equation fonts are a minefield
OMML math runs depend on `w:ascii="Cambria Math"`; the moment a document-wide font replacement touches a math region, rendering breaks.
**Countermeasure**: split document.xml on `<m:oMath...>...</m:oMath>`, do the rFonts replacement **only outside**; verify the `Cambria Math` count is identical before and after.

### 2.3 Detect display equations by `<m:oMathPara`, not by "no text after tag-stripping"
Display equations are wrapped in `<m:oMathPara>`; the "empty plain text" test fails — the content of `\text{(n)}` stays in `m:t`, leaving `(n)` behind after stripping. Inline equations are bare `<m:oMath>` — exactly the discriminator.

### 2.4 Caption detection must use adjacency, not text prefixes
A body sentence like "Table 6 compares this material point…" also matches `^(Figure|Table)\s+\d+`. Correct criteria:
- figure caption = starts with `Figure N` AND the **previous non-empty block** is an image paragraph (contains `<w:drawing` and has no text of its own);
- table caption = starts with `Table N` AND the **next non-empty block** is a `<w:tbl>`.

### 2.5 Where `w:jc` goes inside pPr
Schema order: pStyle → … → spacing/ind → **jc** → … → rPr. Inserting jc at the head of `<w:pPr>` violates the order and Word may report "unreadable content". Safe method: if pPr contains `<w:rPr>`, insert before it; otherwise append at the end of pPr.

### 2.6 A docx is just a zip: edit the XML with Python zipfile
No python-docx needed: read every entry → modify `word/document.xml`, `word/styles.xml` → repackage. Everything other than the deleted/edited entries (images etc.) is written back verbatim with writestr.

## 3. Engineering practice

### 3.1 Verify before overwriting (saved us twice)
Flow: **temp copy → verify → os.replace onto the original**. Verification items depend on content, e.g.: image-centering count / table-border count / caption-style count / equation-alignment count / Cambria Math count unchanged / sectPr margins present / pandoc read-back parses. Two real bugs (math font wrongly replaced; equation numbers corrupted) were both stopped by this gate — no user file was ever polluted.

### 3.2 Word file locks
An open Word document leaves a `~$xxx.docx` lock file, and `os.replace` then raises `PermissionError`. Countermeasure: catch it, write to `xxx-new.docx` instead, and tell the user to close Word — never fail silently or force-delete the lock.

### 3.3 Regenerate-from-md vs patch-in-place — you need both paths
- The user **hand-edited the Word file** → in-place patch only (`_fix_docx_format.py`); regeneration loses the edits;
- The user wants a **fresh build from md** → full pipeline (`_rebuild_word.py`).
Ask first, or check the `~$` lock file to see whether the document is being edited.

### 3.4 Inline Python one-liners inside PowerShell are quote hell
Python code inside a double-quoted PS string gets eaten by the PS parser (`[`, `<`, quote escaping) with bizarre errors. Anything beyond one or two lines goes into a `.py` file, always.

### 3.5 The reference-doc dissection playbook (copying the format of a docx you like)
With zipfile, read `word/document.xml` + `word/styles.xml` + `word/theme/theme1.xml` and inspect in order:
1. `<w:sectPr>`: paper, margins;
2. `<w:docDefaults>` + the Normal style: default font and size (note: theme fonts require digging into theme1.xml's minorFont);
3. Survey the first 40 body paragraphs: each one's style/jc/ind/spacing/sz + a text preview → line spacing, indents, headings and caption conventions fall right out;
4. Run-level `w:rFonts` statistics (a Counter) → the fonts actually used (one case: 1243 TNR runs, while docDefaults' theme font was a decoy);
5. Caption placement (before/after the figure) and equation numbering style (oMathPara jc; number inside the math or in text).

## 4. Second round of pitfalls (equation layout + caption rules)

### 4.1 String concatenation truncates the raw prefix: in `r'...\1' + X + '...\2'` the second segment becomes a plain string ⚠️
When the replacement string is built by concatenation, only the first segment carries the `r` prefix: `r'\1<w:rPr>' + CAP_SZ + '</w:rPr>\2'` — the trailing `\2` is octal-escaped by the plain string into a `\x02` control character, destroying XML well-formedness. Same family as 1.3's `$1`/`\1` but sneakier: **every feature-count verification still passes** (the `<m:r><w:rPr>` prefix matches fine); only XML parsing exposes it. Every concatenated segment must be `r''`.

### 4.2 OMML math-run sizing is a separate mechanism
A math run's style attributes live in **`m:rPr`** (`m:sty`, `m:nor`, …), not `w:rPr`; pandoc-generated `m:r` often lacks `w:rPr` entirely → math inside captions does not inherit 9 pt. Fix: insert `<w:rPr><w:sz …/></w:rPr>` into every `<m:r>` inside captions (schema requires w:rPr **before** m:rPr). Lesson: regexes written against `w:rPr` do nothing to math runs — dump the actual XML shape first, then write.

### 4.3 Pandoc self-closing tags carry a space: `<w:b />` ≠ `<w:b/>`
pandoc outputs `<w:b />` (space before the slash); your own injections write `<w:b/>`. Verification regexes must be `<w:b\s*/>`, otherwise they count only your own injections and none of pandoc's (one case: body bolding all correct, verification reported 0).

### 4.4 Display-equation tab-stop layout requires stripping `oMathPara`
The Word idiom "equation centered + number flush right" is paragraph tab stops (center 4153 / right 8306) + `[tab]equation[tab](n)`, but pandoc wraps display equations in an `<m:oMathPara>` display block that conflicts with runs on the same line. The reference manuscript's equation paragraphs are bare `<m:oMath>` + text runs. Countermeasure: strip the wrapper (together with `oMathParaPr`) and extract the number from the tail of the OMML generated by `\qquad\text{(n)}` (the run with `m:nor` + `m:sty="p"`) into a text run.

### 4.5 `|=` on an outer set inside a Python closure raises UnboundLocalError
Inside a nested function, `seen |= nums` is an assignment to the name `seen` → it becomes local for the whole function and the first read errors out. Use `seen.update(nums)` (a method call, not an assignment).

### 4.6 Verification must include XML well-formedness + pandoc read-back
Feature counts alone (counts, centering, font tags) missed 4.1's control-character pollution. The standard pair: `xml.dom.minidom.parseString()` on document.xml/styles.xml + read-back via `pandoc -t plain`. With both gates in place, problems are caught before shipping.

### 4.7 md caption lines with bold prefixes don't match `^(Figure|Table)`
Source md whose captions already carry bold prefixes (`**Figure 2** …`) get missed by skip rules written against line-initial `^Figure`; use a pattern that tolerates emphasis prefixes like `^\*{0,3}\s*(Figure|Table)\b`, otherwise the body-bolding pass wraps the caption in another layer of `****`.

## 5. Deliverables (shipped in this skill)

| File | Description |
|---|---|
| `references/word-format-spec.md` | The extracted format specification (full tables: page/typography/headings/captions/equations/tables) |
| `scripts/_rebuild_word.py` | Full md→docx pipeline: `\tag` preprocessing, style injection, post-processing, verify-before-write |
| `scripts/_fix_docx_format.py` | In-place patches on hand-edited docx: TNR, image centering, table gridlines |
