# CCD - 软著源代码文档生成器

> Code Copyright Document Generator - 一键生成符合软著要求的源代码文档

## 📖 简介

CCD 是一个专为软件著作权申请设计的源代码文档生成工具。它能够自动扫描项目源码，按照软著规范生成符合格式要求的 Word 文档（.docx），大幅简化软著材料准备工作。

### 核心特性

- 🎯 **完全符合软著规范** - 自动处理页眉、页脚、页码、行距等格式要求
- 🚀 **图形化界面** - 可视化操作，无需记忆命令行参数
- 📦 **智能代码识别** - 支持 30+ 种编程语言，自动识别注释和空行
- 🔧 **灵活配置** - 支持自定义字体、排版、页码格式
- ⚡ **60页模式** - 自动提取前30页+后30页，符合软著要求
- 📊 **代码统计** - 实时统计代码行数，辅助决策
- 🎨 **文件预览** - 支持文件选择和代码预览
- 📝 **日志记录** - 完整的操作日志，便于问题排查

## 📦 安装

### 环境要求

- Python 3.6+
- Windows / macOS / Linux

### 安装依赖

```bash
pip install -r requirements.txt
```

### 依赖包说明

```
click           # 命令行接口
python-docx     # Word 文档生成
PyQt5           # 图形界面框架
qfluentwidgets  # 现代化UI组件库
```

## 🚀 快速开始

### 方式一：图形界面（推荐）

```bash
python cli.py --gui
```

### 方式二：命令行

```bash
# 基础用法
python cli.py -i ./src -e py -o code.docx

# 完整示例
python cli.py \
  --title "我的软件V1.0" \
  --indir ./src \
  --indir ./lib \
  --ext py \
  --ext java \
  --exclude ./tests \
  --exclude ./node_modules \
  --outfile output.docx \
  --font-name Consolas \
  --font-size 10.5
```

## 📘 使用指南

### 图形界面使用流程

1. **配置源码目录**
   - 点击"添加目录"选择项目根目录
   - 可添加多个目录

2. **选择文件类型**
   - 点击"选择"按钮
   - 勾选需要的编程语言后缀（如 py, java, js）
   - 系统会自动扫描目录中的所有文件类型

3. **排除无关文件**
   - 添加 `node_modules`、`dist`、`build` 等目录到排除列表
   - 支持读取 `.gitignore` 文件
   - 避免将构建产物和第三方库写入软著文档

4. **扫描文件**
   - 点击"扫描文件数"
   - 查看符合条件的文件数量

5. **统计代码行数**
   - 点击"统计代码行数"
   - 查看过滤后的总行数、代码行、注释行、空行
   - 根据行数决定是否启用60页模式

6. **选择文件（可选）**
   - 点击"选择文件"打开文件选择对话框
   - **左右两栏设计**：左侧可选文件，右侧已选文件
   - **多选支持**：Ctrl+点击多选，Shift+点击范围选择
   - **双击操作**：左栏双击添加，右栏双击移除
   - **顺序调整**：置顶、上移、下移、置底
   - **文件预览**：点击文件名查看代码内容
   - **搜索过滤**：快速定位文件

7. **配置文档格式**
   - 设置页眉标题（软件名称+版本号）
   - 选择字体（推荐 Consolas 等宽字体）
   - 选择页码格式
   - 如果代码 ≥ 3000行，勾选"60页模式"

8. **生成文档**
   - 点击"生成文档"
   - 等待生成完成
   - 点击"打开输出目录"查看结果

### 命令行参数说明

```bash
python cli.py --help
```

| 参数 | 说明 | 示例 |
|------|------|------|
| `-t, --title` | 页眉标题 | `--title "我的软件V1.0"` |
| `-i, --indir` | 源码目录（可多个） | `-i ./src -i ./lib` |
| `-e, --ext` | 文件后缀（可多个） | `-e py -e java -e js` |
| `-c, --comment-char` | 注释前缀（可多个） | `-c "#" -c "//"` |
| `--font-name` | 字体名称 | `--font-name Consolas` |
| `--font-size` | 字号 | `--font-size 10.5` |
| `--space-before` | 段前间距 | `--space-before 0` |
| `--space-after` | 段后间距 | `--space-after 2.3` |
| `--line-spacing` | 行距 | `--line-spacing 10.5` |
| `--exclude` | 排除路径（可多个） | `--exclude ./tests` |
| `-o, --outfile` | 输出文件 | `-o code.docx` |
| `--template` | 模板文件 | `--template template.docx` |
| `--encoding` | 文件编码 | `--encoding utf-8` |
| `--keep-blank-lines` | 保留空行 | - |
| `--keep-comment-lines` | 保留注释 | - |
| `--gui` | 启动图形界面 | - |
| `-v, --verbose` | 调试信息 | - |

## 🎨 高级功能

### 60页模式详解

**工作原理：**
- 自动计算代码总行数
- 当总行数 ≥ 3000行时：
  - 提取前 1500 行（约前30页）
  - 提取后 1500 行（约后30页）
  - 跳过中间部分
- 最终生成约 60 页的文档

**使用场景：**
- 大型项目（≥ 3000行代码）
- 符合软著规定（≥3000行只需提交前30页+后30页）

**启用方式：**
- GUI：勾选"启用60页模式"
- CLI：暂不支持（建议使用GUI）

### 文件选择与预览

**功能特点：**
- 📂 左右两栏布局，操作直观
- 🔍 实时搜索，快速定位文件
- 👁️ 代码预览，支持语法识别
- 📊 显示文件大小、行数、语言类型
- ⚙️ 可调整预览行数（20/50/100/200）

**使用技巧：**
- 先扫描获取所有文件
- 通过搜索框快速过滤
- 点击文件名查看内容
- 调整文件顺序控制文档结构

### 注释过滤

**支持的语言：**
- Python (`#`, `"""`, `'''`)
- JavaScript/TypeScript (`//`, `/**/`)
- Java/C/C++ (`//`, `/**/`)
- Go/Rust/Swift (`//`, `/**/`)
- SQL (`--`, `/**/`)
- HTML/XML (`<!-- -->`)
- CSS (`/**/`)
- Shell (`#`)
- Ruby/Perl (`#`, `=begin =end`)
- 等 30+ 种语言

**智能识别：**
- 自动识别字符串中的注释符号
- 区分单行注释和多行注释
- 保留代码中的注释符号（如 URL 中的 `//`）

### 日志查看

**日志位置：**
- 项目根目录下的 `logs/` 文件夹
- `ccd.log` - 全部日志
- `ccd_error.log` - 错误日志

**日志功能：**
- 查看实时日志
- 切换全部日志/错误日志
- 清空当前日志
- 打开日志目录
- 自动轮转（最大10MB，保留5个备份）

## 🏗️ 项目结构

```
CCD/
├── cli.py              # 命令行入口
├── logger.py           # 日志管理
├── code_scanner.py     # 代码扫描（CodeFinder类）
├── code_processor.py   # 代码处理工具
├── doc_writer.py       # 文档写入（CodeWriter类）
├── doc_generator.py    # 文档生成主流程
├── ui_main.py          # 主窗口界面
├── ui_dialogs.py       # 对话框组件
├── ui_tasks.py         # 后台任务
├── requirements.txt    # 依赖包
├── README.md           # 项目说明
└── logs/               # 日志目录
    ├── ccd.log
    └── ccd_error.log
```

## 📄 许可证

本项目基于 [MIT License](LICENSE) 开源。

## 🙏 致谢

- [python-docx](https://python-docx.readthedocs.io/) - Word文档生成
- [PyQt5](https://www.riverbankcomputing.com/software/pyqt/) - 图形界面框架
- [qfluentwidgets](https://qfluentwidgets.com/) - 现代化UI组件

<div align="center">
  <p>如果这个项目对你有帮助，请给个 ⭐ Star 支持一下！</p>
</div>
