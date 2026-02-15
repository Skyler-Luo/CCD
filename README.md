# 🧾 CCD（Code-Copyright-Docgen）
一个用于生成软著源代码文档（DOCX）的工具

## ✨ 简介
CCD 会扫描源码目录，将代码按文件组织并排版输出为 DOCX，满足软著材料中“源代码文档”的常见格式需求。支持命令行与图形界面，支持多目录/多后缀、排除路径、过滤空行与注释、以及使用 DOCX 模板统一样式。

## ✅ 特性
- 多目录扫描：可指定多个源码目录
- 多语言后缀：按后缀识别语言并过滤注释（如 py/js/ts/go/java/c/cpp 等）
- 过滤规则：默认过滤空行与注释行，也可选择保留
- 排除路径：支持排除文件/目录；GUI 支持读取 .gitignore（可解析的静态路径）
- 文档排版：页眉标题、字体、字号、段前/段后/行距可配置
- 模板支持：可传入 DOCX 模板统一样式

## 🚀 快速开始
```bash
# 1) 安装依赖
pip install -r requirements.txt
# 2) 启动图形界面
python cli.py --gui
```

## 📦 安装
```bash
pip install -r requirements.txt
```

## 🧩 使用
命令行模式：
```bash
python cli.py --help
python cli.py -i ./src -o ./code.docx
```

常用示例：
```bash
# 多目录 + 多后缀
python cli.py -i ./src -i ./tests -e py -e js -e ts -o ./code.docx

# 保留空行/注释（默认会过滤空行与注释）
python cli.py -i ./src -o ./code.docx --keep-blank-lines --keep-comment-lines

# 排除路径（可重复指定）
python cli.py -i ./src --exclude ./src/vendor --exclude ./src/generated -o ./code.docx
```

图形界面：
```bash
python cli.py --gui
```

## 📝 使用建议
- 后缀选择：仅勾选/传入需要纳入文档的语言后缀，避免把构建产物或依赖代码写入软著材料。
- 排除规则：优先通过 `--exclude` 精确排除 `vendor/`、`dist/`、`build/`、`node_modules/` 等目录。
- 样式统一：如有固定格式要求，准备一份 DOCX 模板并在生成时传入以统一字体与段落样式。

## 📄 许可证
本项目基于 MIT License 许可证开源。详情请参阅 [LICENSE](LICENSE) 文件。

---

<div align="center">
  <p>Made with ❤️ by Skyler-Luo</p>
</div>