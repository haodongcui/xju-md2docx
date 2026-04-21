# xju-md2docx

新疆大学本科毕业论文（设计）`Markdown -> DOCX` 导出工具。

这个仓库尽量保持小而直接：

- 根目录保留一个调用脚本和主要文档
- `example/` 放最小示例
- 真正的转换工具本体封装在 `xju_md2docx/`

## 目录结构

```text
xju-md2docx/
├── README.md
├── GUIDE.md
├── LICENSE
├── requirements.txt
├── xju_md2docx.py
├── demo.sh
├── example/
│   ├── thesis-demo.md
│   └── img/
└── xju_md2docx/
    ├── main.py
    ├── resources/
    ├── world-math/
    └── official-materials/
```

## 快速开始

安装 Python 依赖：

```bash
pip install -r requirements.txt
```

导出你自己的论文：

```bash
python xju_md2docx.py thesis.md thesis.docx
```

如果不写输出文件，默认生成同名 `.docx`：

```bash
python xju_md2docx.py thesis.md
```

如果你明确想关闭公式转换：

```bash
python xju_md2docx.py thesis.md thesis.docx --no-formula-conversion
```

运行示例：

```bash
bash demo.sh
```

## 默认行为

- 默认模板：`xju_md2docx/resources/xju-template.docx`
- 默认封面资源：`xju_md2docx/resources/`
- 默认输出：输入文件同名 `.docx`

常用参数：

```bash
python xju_md2docx.py thesis.md thesis.docx --template my-template.docx
python xju_md2docx.py thesis.md thesis.docx --no-template
python xju_md2docx.py thesis.md thesis.docx --no-cover-assets
python xju_md2docx.py thesis.md thesis.docx --no-formula-conversion
```

## 公式支持

如果你希望把 LaTeX 公式尽量转成 Word 原生公式：

```bash
cd xju_md2docx/world-math
npm install
cd ../..
```

默认情况下，脚本会优先尝试把 LaTeX 公式转换成 Word 原生公式。

如果本地没有 `npm` / `node`，或者依赖尚未安装成功，脚本不会自动联网安装依赖，而是会继续导出、明确打印警告，并把未转换成功的公式保留为 LaTeX 文本。

## 写作约定与工作流

详细说明见 [GUIDE.md](GUIDE.md)。

核心建议：

- Markdown 作为唯一长期主稿。
- 内容层面的改动回 Markdown 全量重导。
- Word / WPS 只负责最后的格式验收和少量微调。

## 示例与附加材料

- 示例主稿：[thesis-demo.md](example/thesis-demo.md)
- 默认模板：[xju-template.docx](xju_md2docx/resources/xju-template.docx)
- 备查材料目录：[official-materials](xju_md2docx/official-materials)

说明：

- `xju_md2docx/official-materials/` 只是备查，不参与核心运行。
- 学校模板和规范可能按年份、学院或教务要求调整。
- 真正提交前，请以你所在学院当前最新要求为准。
- 如果你准备公开分发这些学校材料，请自行确认分发政策。

## 已知限制

- 最稳妥的模板模式依赖一个现成的 `docx` 模板。
- 复杂嵌套列表、脚注、自动编号公式、跨页表格微调，目前没有做成完整的 Word 排版系统。
- 这个工具解决的是“主稿统一维护”和“重复导出”，不是完全替代最后的人工检查。

## License

代码部分采用 [MIT License](LICENSE)。
