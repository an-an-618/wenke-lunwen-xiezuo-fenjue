# Zotero 与 Word 写作规范

## Zotero 工作流

1. 检查 Zotero Desktop 或本地 API 是否可用。
2. 新建或选定 collection，名称按用户要求。
3. 导入题录、PDF、网页或 BibTeX。
4. 给条目添加标签：已读、可引用、待核验、理论背景、原始文本。
5. 导出 BibTeX、CSL JSON 或按需插入 Word 引用。

导入后不要默认认为文献可引用。只有在读过全文或可靠摘录后，才可把它用于正文观点。

## Word 文档交付

创建 DOCX 时注意：

- 标题、摘要、关键词、正文、参考文献层级清楚。
- 不自动添加不必要页头。
- 不使用蓝色标题，除非用户要求。
- 不插入装饰性图表。
- 引用格式与参考文献格式统一。
- 中文标点、书名号、引号保持一致。

## APA 与其他格式

APA 适合社会科学与跨学科写作，中文文学论文有时使用脚注或 GB/T 7714。应按用户指定格式执行。若用户没有指定：

- 课程论文或用户明确要求 APA：使用 APA 文内引用。
- 中文核心期刊模拟：可询问是否使用脚注或 GB/T 7714；若不询问，保持用户前文已指定格式。
- 外文理论文献：保留原名或通行中译名，首次出现可括注外文。

## Zotero 与 Word 的协作边界

智能体可以帮助：

- 创建 collection。
- 导入或导出题录。
- 生成参考文献列表。
- 根据 Zotero 数据核对作者、年份、标题。

智能体不能假装：

- 已经通过 Zotero 插件在用户 Word 界面完成不可见操作。
- 已读 Zotero 中没有打开或导出的 PDF。
- 已核验没有证据文件的文献观点。

## 最终核验

交付前执行：

```bash
python scripts/docx_basic_qa.py paper.docx
python scripts/citation_skeleton.py paper.docx --out citation-audit.md
```

再人工或智能体逐条补充核验证据。
