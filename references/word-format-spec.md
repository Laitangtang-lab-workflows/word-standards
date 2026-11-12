# Word manuscript format specification

**Basis**: a published reference manuscript docx from the research group (dissected 2026-09-10; every rule below was measured from it).
**Applies to**: manuscripts generated from md into Word in the reference-manuscript style (md → Word automatic pipeline).

---

## 1. Page setup

| Item | Rule |
|---|---|
| Paper | A4 (11906 × 16838 twips) |
| Top/bottom margins | 1440 twips = 2.54 cm |
| Left/right margins | 1800 twips = 3.17 cm |
| Header/footer distance | 851 / 992 twips |
| Header/footer content | none (the reference manuscript used none) |

## 2. Body text

| Item | Rule |
|---|---|
| Western font | Times New Roman |
| CJK font | SimSun (only where a little Chinese appears) |
| Size | 10.5 pt (五号, sz=21) |
| Line spacing | double (line=480, lineRule=auto) |
| Alignment | justified (jc=both) |
| First-line indent | 210 twips (1 character, firstLineChars=100) |

## 3. Headings

| Level | md level | Rule |
|---|---|---|
| Title | `#` | TNR 12 pt bold, double-spaced |
| Level-1 heading (1. Introduction) | `##` | TNR 12 pt bold, before 240 twips, double-spaced, left-aligned |
| Level-2 heading (2.1 …) | `###` | TNR 10.5 pt bold, before 240 twips, double-spaced, left-aligned |

Numbers are written inside the heading text (`## 1. Introduction`); Word auto-numbering is not used.

## 4. Figures and captions

- **Figures**: own paragraph, horizontally centered. Width rule (added 2026-09-17, `size_images` post-pass): prefer the physical size computed from the image file's own DPI metadata (= pixels ÷ DPI, i.e. the size set in Illustrator); if the physical width exceeds the text block (14.65 cm), scale proportionally down to fit; **images without DPI metadata fall back to pandoc's default behavior (96 dpi assumption → full column width)**. So when exporting from Illustrator always use "Export As" and fill in an explicit ppi (e.g. 300) so the physical size lands in the file; "Save for Web / Export for Screens" discards it. In md you can also force a width with `![](path){width=8cm}` (pandoc-native, takes priority over DPI).
- **Captions**: one paragraph below the figure, format `Figure N. text`; body 9 pt (sz=18), no first-line indent, justified.
- **Bold caption labels**: the leading `Figure N.` (or `Figure N`, number plus the following period) in captions is **bold**, the rest not; `Table N.` labels likewise.
- **Inline math in captions**: must match the caption body size (9 pt). OMML math runs do not inherit caption sizes natively; the pipeline force-inserts `w:sz=18` into every `m:r` inside captions.
- **First occurrence in body bold**: the **first** reference to each figure number in body text (`Fig. N` / `Figure N` / `Figs. N–M`) is bolded; later occurrences are not; a range `Figs. 4–6` counts as first occurrence of 4, 5 and 6. Table references are not bolded for now (enable if needed).
- md style: the image line's alt stays empty (`![](./figures/xxx.png)`); the caption is a separate plain paragraph hand-written as `Figure N. ...`; the md needs **no** manual bolding anywhere (label bolding and first-occurrence bolding are both handled by the pipeline).
  The pipeline already disables implicit figures with `-f markdown-implicit_figures`, so alt text never becomes a duplicate caption.

## 5. Tables and table captions

- **Table captions**: one paragraph above the table, format `Table N. text`, same style as figure captions (9 pt, no indent).
- **Tables**: full 0.5 pt single-line grid (tblBorders six sides single sz=4), cell margins 30 top/bottom / 80 left/right twips, text TNR 10.5 pt.
- The reference manuscript itself has no tables; the grid rule is this project's own addition.

## 6. Equations

- Display equations: Word-native OMML (Cambria Math); **equation centered, number `(n)` flush right** — implemented with paragraph tab stops (center tab 4153 twips, right tab 8306 twips), equation and number separated by a tab, paragraph left-aligned. The pipeline strips pandoc's `oMathPara` wrapper (down to bare `m:oMath`, matching the reference manuscript) and extracts the number from inside the math into a TNR 10.5 pt text run at the right tab.
- The md source keeps the standard LaTeX `\tag{n}`; the pipeline performs "number extraction + tab-stop layout" automatically at generation time (pandoc silently drops `\tag`, so it can never be relied on directly).
- Inline math `$...$` follows the body text, no special handling (captions are the exception, see §4).
- Equation fonts must stay Cambria Math; any document-wide font replacement skips `<m:oMath>` regions.

## 7. Generation pipeline

```
# Scripts ship with the word-standards skill (skill-path\scripts\); run from the manuscript directory
python <skill-path>\scripts\_rebuild_word.py input.md output.docx
```

The pipeline performs: `\tag` preprocessing → ```reference fenced block → `## References` heading + entry paragraphs → styling on top of pandoc's default reference.docx (every rule above) → pandoc conversion (implicit figures off) → post-processing (captions 9 pt, images centered, **image widths from DPI metadata**, equation right-tab layout, table gridlines, **EndNoteBibliography style on reference entries**) → verify-before-write → overwrite the target (if Word holds the file, output goes to `-new.docx`).

**Caution**: regenerating straight from md loses manual edits made in Word. If the file was only edited in Word and needs in-place formatting (the three items TNR / image centering / table gridlines), use instead:

```
python <skill-path>\scripts\_fix_docx_format.py target.docx
```

## 8. Verification checklist (auto-checked after every build)

The counts below are a snapshot from one real manuscript's build.

- [x] XML well-formed (minidom parses) + pandoc read-back parses
- [x] all 9 images centered
- [x] all 8 tables gridded
- [x] all 17 figure/table captions at 9 pt (9 figure + 8 table; body sentences like "Table N compares…" must not be touched)
- [x] all 17 `Figure N.` / `Table N.` labels bold
- [x] every inline-math run inside captions at 9 pt
- [x] 20 display equations: 17 numbers extracted to the right tab, `oMathPara` wrappers reduced to zero
- [x] first in-text figure/table references bolded (7 in that manuscript: Fig. 1, Fig. 2, Fig. 3, Figs. 4–6, Fig. 7, Figure 8, Fig. 9)
- [x] Cambria Math font damage: zero
- [x] A4 margins correct (left=1800, top=1440)
- [x] References: 57 entries each an own paragraph with the EndNoteBibliography style (10 pt / double-spaced / no indent), no code-block residue

## 9. References

> Extracted from the reference manuscript's References section (parsed 2026-09-12; its entry style is EndNoteBibliography — TNR 10 pt, double-spaced, no first-line indent).

| Item | Rule |
|---|---|
| Heading paragraph | text `References`, treated as `##` (TNR 12 pt bold, before 240, double-spaced — same as §3 level-1 headings) |
| Entry paragraph style | EndNoteBibliography: TNR **10 pt** (sz=20), double-spaced (line=480 auto), no first-line indent, justified |
| Entry text | `[N] A. Author, B. Author, Title, J. Abbrev. Vol(Issue) (Year) Pages.` |

Text-format details (entry assembly is carried by the academic-markup skill's `lib/format-reference.mjs`):

- The number `[N]` is followed by a **space** (not a comma); N follows first-occurrence order in the body.
- Authors: surname first, initials after (`L.S. van Hazendonk`, `B.P. Croom` style, initials run together), comma-separated, all listed; >10 authors → first 3 + `et al.`.
- Journals use ISO 4 abbreviations (`Addit. Manuf.`, `J. Non-Newton. Fluid Mech.`, `Bioact. Mater.`).
- Volume/issue/year/pages: `36 (2024) 102014.`; the issue hugs the volume without a space `3(2) (2018)`; without pages, stop at `(2026).`.
- **No DOI/URL output** (the reference manuscript contains none throughout).

md style: the bibliography stays inside the ```reference fenced block (the academic-markup `refs` step's contract format, idempotently renumberable); `_rebuild_word.py` converts the block into `## References` + entry paragraphs at build time (blank lines between entries inserted by the script) and applies the EndNoteBibliography style to the entries.
