# 格式大师

将 Markdown 和纯文本转换为可编辑的 Word 文档。

支持中文排版、标题、多级编号、表格、图片和代码块。转换时会自动跳过文首的 YAML 笔记属性。

## 在 Windows 上运行源码

需要 Python 3.13。在 PowerShell 中进入项目目录，再运行：

```powershell
py -3.13 -m venv .venv
& '.\.venv\Scripts\python.exe' -m pip install -r '.\requirements.txt'
& '.\.venv\Scripts\python.exe' '.\scripts\convert.py' --input '.\verification\sample.md' --output '.\sample.docx' --offline --no-embed-fonts
```

生成的 `sample.docx` 会保存在项目目录。源码版使用电脑上已有的字体。