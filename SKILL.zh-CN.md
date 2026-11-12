---
name: word-standards
description: >-
  Word 文档格式规范与生成管线（md → pandoc → docx + OOXML 后处理）。当用户要"生成 Word""转 docx""按课题组格式出 Word""pandoc 转换""论文 Word 版式""公式编号""图注表注格式"或要解剖参考稿 docx 抄格式、对已生成的 docx 原地补格式时触发。核心坑：\tag{} 被 pandoc 静默丢弃、隐式 figure 重复图注、Python/PowerShell 反向引用语法差（$1 vs \1）、OMML 数学字体雷区、写前必须验证（临时副本→验证→os.replace）。英文正本见 SKILL.md。
---

# Word Docgen — md→pandoc→docx 生成管线（中文伴侣件）

> 本文件为中文速览，英文正本（内容一致、以它为准）见 [SKILL.md](SKILL.md)，深读材料 references/ 仅有英文版。

中文 md（或翻译后的英文 md）按课题组参考稿格式生成 Word：pandoc 转换 + Python zipfile 直接读写 docx 的 XML 后处理。

## 六条总铁律（先记）

1. **`\tag{n}` 会被 pandoc 静默丢弃**：生成前把 `\tag{…}` 替换为 `\qquad\text{(n)}`，编号进公式内部右端；md 源保持标准 `\tag` 写法不动。
2. **关掉隐式 figure**：`-f markdown-implicit_figures`，图片 alt 留空——否则 pandoc 以 alt 文本自动生成重复图注。
3. **反向引用语法跨语言要命**：Python `re.sub` 用 `\1`/`\g<1>`，PowerShell/JS 用 `$1`；带错习惯输出字面量 `($1)`。
4. **公式字体是雷区**：OMML 数学 run 依赖 `w:ascii="Cambria Math"`；全文字体替换必须按 `<m:oMath>…</m:oMath>` 切分、只对区外做，验证 Cambria Math 计数前后相等。
5. **写前必须验证**：临时副本 → 验证（特征计数 + XML 良构性 `minidom.parseString` + pandoc 回读）→ `os.replace` 覆盖；两次真实 bug 都是验证拦下的。
6. **改样式表优于改正文 XML**：pandoc 从 `--reference-doc` 继承样式与 sectPr；解剖参考稿的套路见 references §3.5。

## 可复用脚本（skill 自带 `scripts\`，随 skill 分发）

- `_rebuild_word.py`：md→docx 全管线（`\tag` 预处理、样式注入、后处理、写前验证）；
- `_fix_docx_format.py`：对已手工编辑过的 docx 原地补三项格式（TNR、图片居中、表格网格线）；
- 调用（在手稿目录下运行，输入输出为相对路径）：`python <skill 路径>\scripts\_rebuild_word.py 输入.md 输出.docx`。

选择依据：用户手工改过 Word → 只能原地补丁（重生成会丢修改）；要求从 md 重出 → 全管线。先问清或看 `~$` 锁文件判断文档是否在编辑中。

## 分工边界

- Word 批注/修订（docx-editor MCP 的场景）：不归本 skill；
- 手稿版式规范本身（页面/字体/标题惯例）：本 skill `references/word-format-spec.md` 自带；
- 驱动脚本的 PowerShell 编码坑（引号转义链）：见 references/word-pandoc-pitfalls.md §3.4。
