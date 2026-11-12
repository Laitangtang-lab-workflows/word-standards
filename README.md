# word-standards

**Generate journal-style Word documents from Markdown: a pandoc + raw-OOXML pipeline with every trap already defused — equation numbers, caption rules, math fonts and all.**

> [!NOTE]
> This repository is **written and maintained entirely by AI agents** under human direction — see [AI-maintained](#ai-maintained).
> Documentation is English-first; a Chinese README is available ([README.zh-CN.md](README.zh-CN.md)), and a Chinese SKILL companion ([SKILL.zh-CN.md](SKILL.zh-CN.md)).

## Why

Turning a Markdown manuscript into a properly formatted .docx is deceptively hard: pandoc silently drops `\tag{}` equation numbers, implicit figures duplicate your captions, one careless font replacement destroys every equation, and the difference between `$1` and `\1` quietly corrupts the XML. This project packages a complete, battle-tested pipeline — distilled from a real research group's manuscript workflow — where every rule exists because a specific bug once slipped through:

- **Full md→docx pipeline** (`_rebuild_word.py`): `\tag` preprocessing, style injection via pandoc's reference.docx mechanism, post-processing (caption sizing, image centering, DPI-aware image widths, tab-stop equation layout, table gridlines, reference-entry styling), and verify-before-overwrite gates.
- **In-place patcher** (`_fix_docx_format.py`): three formatting fixes applied directly to a hand-edited docx, without regenerating and losing the edits.
- **A measured format specification**: page geometry, typography, heading conventions, caption rules, equation layout and reference formatting — every value extracted from a real published manuscript, in copy-pasteable twips/pt tables.
- **The pitfall collection**: twenty-plus confirmed traps across the pandoc layer, the OOXML/regex layer and general engineering practice, each with its countermeasure.

## The skill format

The core of this repository is a **`SKILL.md` agent skill** (the convention used by Claude Code, DeepSeek Harness and other agent harnesses). Drop the repo into your harness's skill directory and an AI session picks up the whole methodology with it:

```
git clone https://github.com/Laitangtang-lab-workflows/word-standards ~/.agents/skills/word-standards
```

You can equally read `SKILL.md` and `references/` as plain documentation, or run the scripts directly — no harness required (pandoc 3.x + Python 3.12).

## Usage

```bash
# full build, run from the manuscript directory
python <skill-path>/scripts/_rebuild_word.py input.md output.docx

# in-place patch on a hand-edited docx
python <skill-path>/scripts/_fix_docx_format.py target.docx
```

## Repository layout

| Path | Content |
|---|---|
| `SKILL.md` | Skill entry point: the six iron rules |
| `references/word-format-spec.md` | The measured format specification |
| `references/word-pandoc-pitfalls.md` | The full pitfall collection with countermeasures |
| `scripts/` | The pipeline and the in-place patcher |

## AI-maintained

Every file in this repository — code, documentation, this README — is written and updated by AI agents under human supervision. The pitfalls were collected while AI sessions built real manuscripts end-to-end. Expect commits to be machine-authored; humans review direction and decisions, not lines.

## License

[CC0-1.0](LICENSE) — released to the public domain. Use, fork and vendor it into anything, no attribution required (a nod is always appreciated).

By submitting a pull request you agree that your contributions are dedicated under CC0 as well.
