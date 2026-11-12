---
name: word-standards
description: >-
  Word document format specification and generation pipeline (md → pandoc → docx + OOXML post-processing). Use when generating Word documents, converting to docx, producing journal-manuscript-style Word output, running pandoc conversions, laying out equations with right-aligned numbers, formatting figure/table captions, dissecting a reference docx to copy its format, or patching formatting in place on an already-generated docx. Core pitfalls: \tag{} silently dropped by pandoc, implicit figures duplicating captions, Python/PowerShell backreference syntax differences ($1 vs \1), the OMML math-font minefield, and mandatory verify-before-overwrite (temp copy → verify → os.replace). Chinese companion: SKILL.zh-CN.md.
---

# Word Docgen — the md→pandoc→docx pipeline

Generate Word documents from Chinese md (or translated English md) styled after a research group's reference manuscript: pandoc conversion + Python zipfile post-processing that reads and writes the docx XML directly.

## Six iron rules (memorize first)

1. **`\tag{n}` is silently dropped by pandoc**: before generating, replace `\tag{…}` with `\qquad\text{(n)}` so the number rides inside the equation's right end; the md source keeps the standard `\tag` spelling untouched.
2. **Disable implicit figures**: `-f markdown-implicit_figures`, image alt text left empty — otherwise pandoc auto-generates a duplicate caption from the alt text.
3. **Backreference syntax kills across languages**: Python `re.sub` uses `\1`/`\g<1>`; PowerShell/JS use `$1`. Carrying the wrong habit outputs the literal `($1)`.
4. **Equation fonts are a minefield**: OMML math runs depend on `w:ascii="Cambria Math"`; any document-wide font replacement must split on `<m:oMath>…</m:oMath>` and touch only the outside, verifying the Cambria Math count is unchanged.
5. **Verify before overwriting**: temp copy → verify (feature counts + XML well-formedness via `minidom.parseString` + pandoc read-back) → `os.replace`; two real bugs were both caught by exactly this gate.
6. **Editing the style sheet beats editing body XML**: pandoc inherits styles and sectPr from `--reference-doc`; the reference-doc dissection playbook is in references §3.5.

## Content map

Full detail (pandoc-layer pitfalls, OOXML/regex-layer pitfalls, engineering practice, the second round of equation-layout and caption-rule pitfalls, reusable scripts) lives in [references/word-pandoc-pitfalls.md](references/word-pandoc-pitfalls.md). The extracted format specification itself (page/typography/headings/captions/equations/tables) lives in [references/word-format-spec.md](references/word-format-spec.md).

## Reusable scripts (shipped in `scripts\`)

- `_rebuild_word.py`: the full md→docx pipeline (`\tag` preprocessing, style injection, post-processing, verify-before-write);
- `_fix_docx_format.py`: patches three formatting items in place on a hand-edited docx (TNR, image centering, table gridlines);
- Invocation (run from the manuscript directory; inputs and outputs are relative paths): `python <skill-path>\scripts\_rebuild_word.py input.md output.docx`.

Choosing between them: the user hand-edited the Word file → in-place patch only (regeneration loses edits); the user wants a fresh build from md → full pipeline. Ask first, or check for a `~$` lock file to see whether the document is open for editing.

## Scope boundaries

- Word comments/track changes (the docx-editor MCP territory): not this skill;
- The manuscript format spec itself (page/typography/heading conventions): carried by this skill's `references/word-format-spec.md`;
- PowerShell quoting pitfalls when driving the scripts: see references/word-pandoc-pitfalls.md §3.4.
