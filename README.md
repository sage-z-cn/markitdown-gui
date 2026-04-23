# MarkItDown GUI

基于 [Microsoft MarkItDown](https://github.com/microsoft/markitdown) 的图形界面工具，支持将多种文件格式拖放转换为 Markdown。

## 功能特性

- **拖放转换** — 将文件拖入窗口即可转换为 Markdown
- **批量处理** — 支持同时拖入多个文件
- **格式丰富** — 支持 PDF、Word、PowerPoint、Excel、图片、音频、HTML、EPub 等格式
- **转换历史** — 基于 SQLite 存储的转换记录，支持查看、重试、删除
- **快速定位** — 点击文件名直接打开源文件或转换结果，一键打开所在目录
- **输出配置** — 可选择输出到原始文件目录或自定义目录

## 环境要求

- Python 3.10+
- Windows / macOS / Linux

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/sage-z-cn/markitdown-gui.git
cd markitdown-gui
```

### 2. 创建虚拟环境

```bash
python -m venv venv
# Windows
.\venv\Scripts\Activate.ps1
# macOS / Linux
source venv/bin/activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 运行

```bash
python main.py
```

## 打包为可执行文件

使用 PyInstaller 打包为独立 exe：

```powershell
.\build.ps1
```

或手动执行：

```bash
pyinstaller --noconfirm --onefile --windowed --name "markitdown-gui" --icon "assets/logo.ico" --add-data "assets;assets" --add-data "venv\Lib\site-packages\magika;magika" main.py
```

打包产物位于 `dist/markitdown-gui.exe`。

## 项目结构

```
markitdown-gui/
├── main.py              # 主窗口、UI 交互逻辑
├── converter.py         # 文件转换线程
├── config.py            # 配置读写与输出目录管理
├── database.py          # SQLite 数据库操作
├── config.json          # 运行时配置文件
├── requirements.txt     # Python 依赖
├── build.ps1            # Windows 打包脚本
├── MarkItDown-GUI.spec  # PyInstaller 规格文件
├── assets/
│   └── logo.ico         # 应用图标
└── .docs/
    └── markintdown.md   # MarkItDown 库参考文档
```

## 使用方法

1. 启动应用后，将待转换的文件拖放到窗口中，或点击「选择文件」按钮
2. 转换完成后，进度条显示 100%，状态显示「转换成功」
3. 在历史记录表格中：
   - 点击**源文件名** → 打开原始文件
   - 点击**输出文件名** → 打开 Markdown 文件
   - 点击「重试」→ 重新转换
   - 点击「删除」→ 删除记录及输出文件
   - 点击「目录」→ 打开输出文件所在目录
4. 点击 ⚙「设置」可配置输出目录

## 技术栈

| 组件 | 技术 |
|------|------|
| GUI 框架 | PyQt5 |
| 文件转换 | markitdown |
| 本地存储 | SQLite |
| 打包工具 | PyInstaller |

## 许可证

请参阅 [MarkItDown](https://github.com/microsoft/markitdown) 的许可证信息。
