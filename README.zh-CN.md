# word-standards

**从 Markdown 生成期刊格式 Word：pandoc 加原生 OOXML 管线，公式编号、图注规则、数学字体等坑全部预先排掉。**

> [!NOTE]
> 本仓**全部内容由 AI 维护**，人类定方向、做决策，详见 [AI 维护声明](#ai-维护声明)。英文说明见 [README.md](README.md)。

## 解决什么问题

把 Markdown 手稿转成格式合规的 .docx 出乎意料地难：pandoc 静默丢弃 `\tag{}` 公式编号、隐式 figure 复制你的图注、一次随手全文字体替换毁掉全部公式、`$1` 与 `\1` 的差别悄悄污染 XML。本项目打包了一条完整且经过实战检验的管线——提炼自真实课题组手稿流程——**每条规则背后都有一次具体踩坑**：

- **md→docx 全管线**（`_rebuild_word.py`）：`\tag` 预处理、经 pandoc reference.docx 机制注入样式、后处理（图注字号、图片居中、按 DPI 定图宽、制表位公式版式、表格网格线、文献条目样式）、写前验证闸门
- **原地补丁器**（`_fix_docx_format.py`）：对手工编辑过的 docx 原地补三项格式，不重生成、不丢修改
- **实测格式规范**：页面几何、字体字号、标题惯例、图注规则、公式版式、文献格式——每个数值都从真实发表论文提取，twips 与 pt 可直接抄表
- **踩坑全集**：pandoc 层、OOXML 与正则层、工程实践三层二十余条实锤坑，每条附对策

## skill 形态与安装

本仓核心是一个 **`SKILL.md` 格式的 agent skill**（Claude Code、DeepSeek Harness 等框架通用）。clone 进技能目录，AI 会话即可带着整套管线工作：

```
git clone https://github.com/Laitangtang-lab-workflows/word-standards ~/.agents/skills/word-standards
```

也可以把 `SKILL.md` 与 `references/` 当普通文档读，或直接跑脚本（pandoc 3.x 加 Python 3.12）。

## 使用

```bash
# 全管线构建，在手稿目录下运行
python <skill路径>/scripts/_rebuild_word.py 输入.md 输出.docx

# 对手工编辑过的 docx 原地补格式
python <skill路径>/scripts/_fix_docx_format.py 目标.docx
```

## 仓库结构

| 路径 | 内容 |
|---|---|
| `SKILL.md` | skill 入口：六条总铁律 |
| `references/word-format-spec.md` | 实测格式规范 |
| `references/word-pandoc-pitfalls.md` | 踩坑全集与对策 |
| `scripts/` | 全管线与原地补丁器 |

## AI 维护声明

本仓全部内容（代码、文档、本 README）由 AI 会话在人类监督下编写与更新。踩坑记录收集自 AI 会话端到端构建真实手稿的过程。提交均由机器生成，人类把关方向与决策，不逐行审码。

## 证书

[CC0-1.0](LICENSE)——公有领域贡献。可任意使用、fork、并入任何项目，无须署名（注明出处是善意但非义务）。

提交 PR 即表示同意你的贡献同样以 CC0 贡献。
