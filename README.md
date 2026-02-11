# CCD
Code-Copyright-Docgen：一个软著源代码文档生成器

## 简介
CCD 用于扫描源码目录并生成软著所需的源代码文档（docx），支持命令行与图形界面。

## 安装
```bash
pip install -r requirements.txt
```

## 使用
命令行模式：
```bash
python cli.py --help
python cli.py -i ./src -o ./code.docx
```

图形界面：
```bash
python cli.py --gui
```

## 依赖
- click
- python-docx
- PyQt5
- qfluentwidgets

## 许可证
MIT License
