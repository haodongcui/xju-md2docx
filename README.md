# xju-md2docx

新疆大学本科毕业论文（设计）`Markdown -> DOCX` 导出工具。

这套工具的定位很明确：

- 用 Markdown 长期维护论文主稿
- 用 `docx` 模板继承学校版式骨架
- 导出 Word 供导师审阅、格式检查和最终提交

## 已完成功能

- [x] Markdown 主稿导出为 `docx`
- [x] 继承学校 `docx` 模板的节属性与基础样式
- [x] 自动生成封面、摘要、英文摘要和目录页
- [x] 支持一级到三级标题
- [x] 支持单图与并排图
- [x] 支持 Markdown 管道表格
- [x] 支持正文引用跳转到参考文献
- [x] 支持 LaTeX 公式转换为 Word 原生公式
- [x] 公式依赖缺失时打印清晰 warning 并保底导出
- [x] 提供示例论文与一键 demo 导出脚本

## 安装

### 环境要求

- Python 3
- `pip`
- 如果你希望公式导出为 Word 原生公式：
  - `node`
  - `npm`

### 推荐安装方式

推荐直接把 Python 依赖和公式依赖一次装好。

```bash
cd xju-md2docx
pip install -r requirements.txt
cd xju_md2docx/world-math
npm install
cd ../..
```

这样导出时：

- 普通正文、图片、表格、目录可以正常工作
- LaTeX 公式会尽量转换成 Word 原生公式

### 最小安装方式

如果你当前只想先试用正文导出，也可以只装 Python 依赖：

```bash
cd xju-md2docx
pip install -r requirements.txt
```

这种情况下，导出仍然能完成，但公式会保留为 LaTeX 文本，并打印 warning。

## 快速开始

### 运行示例

```bash
cd xju-md2docx
bash demo.sh
```

示例主稿在 [example/thesis-demo.md](example/thesis-demo.md)。

### 导出你自己的论文

```bash
cd xju-md2docx
python xju_md2docx.py thesis.md thesis.docx
```

如果不写输出文件，默认生成同名 `.docx`：

```bash
python xju_md2docx.py thesis.md
```

## 公式模式说明

默认情况下，脚本会优先尝试把 LaTeX 公式转换成 Word 原生公式。

如果公式依赖没有装好，脚本不会自动联网安装，而是：

- 继续完成导出
- 明确打印 warning
- 把未转换成功的公式保留为 LaTeX 文本

如果你明确不想启用公式转换：

```bash
python xju_md2docx.py thesis.md thesis.docx --no-formula-conversion
```

## 常用命令

```bash
python xju_md2docx.py thesis.md thesis.docx --template my-template.docx
python xju_md2docx.py thesis.md thesis.docx --no-template
python xju_md2docx.py thesis.md thesis.docx --no-cover-assets
python xju_md2docx.py thesis.md thesis.docx --no-formula-conversion
```

默认行为：

- 默认模板：`xju_md2docx/resources/xju-template.docx`
- 默认封面资源：`xju_md2docx/resources/`
- 默认输出：输入文件同名 `.docx`

## Markdown 主稿怎么写

推荐主稿结构：

```markdown
# 论文题目

## 封面信息
论文题目：你的论文题目
学生姓名：张三
学号：2022xxxxxx
所属院系：某某学院
专业：某某专业
班级：某某班
指导教师：某某老师
日期：2026 年 4 月

---

## 声明
这里写声明正文。

---

## 摘要
这里写中文摘要。

关键词：关键词1；关键词2；关键词3

---

## ABSTRACT
Here is the English abstract.

KEY WORDS: Keyword 1; Keyword 2; Keyword 3

---

## 目录
这里的文字只作为占位提示，最终目录由 Word 域生成。

---

# 1 绪论
## 1.1 研究背景
正文。

# 参考文献
[1] ...

# 致谢
...

# 附录
## 附录 A ...
...
```

注意：

- 正文必须从 `# 1 ...` 这种编号一级标题开始
- 正文最多稳定支持到三级标题
- `参考文献`、`致谢`、`附录` 按不编号标题处理
- `附录` 下的二级标题也按不编号标题处理

### 封面信息字段

推荐字段：

```text
论文题目
学生姓名
学号
所属院系
专业
班级
指导教师
日期
```

## 段落、公式、引用

普通段落：

- 用空行分段
- 不要为了视觉换行而频繁手动断句
- 同一段里可以混用中文、英文、行内公式和引文

行内公式：

```markdown
模型目标写作 $J(\theta)$。
```

块公式：

```markdown
$$
J(\theta)=\mathbb{E}\left[\sum_{t=0}^{T}\gamma^t r_t\right]
$$
```

参考文献：

```markdown
# 参考文献

[1] Author A. Title[J]. Journal, 2024.
[2] Author B. Title[C]//Conference. 2025.
```

正文引用：

```markdown
相关结论可参考 Hafner 等人的工作[1-2]。
```

说明：

- 每条参考文献必须以 `[数字]` 开头
- 正文中的 `[1]`、`[1-2]`、`[1, 3]` 会尽量链接到参考文献位置

## 图片与表格

单图：

```markdown
![图 2-1 示例图](img/example.png)

图 2-1 示例图
```

并排图：

```markdown
:::figure-row
![(a) 方法 A](img/a.png)
![(b) 方法 B](img/b.png)
:::

图 2-2 两种方法对比
```

表格：

```markdown
表 2-1 示例表

| 指标 | 数值 |
| --- | --- |
| 准确率 | 91.2% |
| 召回率 | 88.5% |
```

建议：

- 图和图题分开写
- 表题写在表格前
- 并排图一行放 2 张左右最稳
- 图片文件名尽量稳定
- 不要依赖复杂合并单元格

## 导出前准备

开始导出前，建议先确认：

- 主稿标题结构已经稳定
- 图片路径都能从 Markdown 文件相对访问到
- 参考文献章节已经写在文末
- 如果你要 Word 原生公式，已经按上面的安装步骤装好公式依赖

如果导出时看到公式 warning：

- 先不要慌，正文仍然会正常导出
- 只是公式被保留成 LaTeX 文本
- 补装公式依赖后重新导出即可

## 导出后检查

- 打开 Word / WPS
- 刷新目录、页码和交叉引用
- 检查封面、中英文摘要、目录分页、图号表号、参考文献格式、附录编号
- 不要把手改后的 `docx` 当成长期主稿

## 仓库结构

```text
xju-md2docx/
├── README.md
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

说明：

- 根目录主要放入口脚本、说明文档和示例
- `xju_md2docx/` 是真正可迁移的工具本体
- `official-materials/` 只是备查，不参与核心运行

## 模板与附件

- 默认模板：[xju-template.docx](xju_md2docx/resources/xju-template.docx)
- 示例主稿：[thesis-demo.md](example/thesis-demo.md)
- 备查材料目录：[official-materials](xju_md2docx/official-materials)

注意：

- 学校模板和规范可能按年份、学院或教务要求调整
- 真正提交前，请以你所在学院当前最新要求为准
- 如果你准备公开分发这些学校材料，请自行确认分发政策

## License

代码部分采用 [MIT License](LICENSE)。
