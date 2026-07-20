# OneConvertXML

将 OneNote 分区文件（`.one`）按页面层级转换为分层 Markdown，针对嵌套表格深度优化。

`.one → 分层 XML → 分层 Markdown`

## 功能特性

- `.one` 转分层 XML（按页面层级建目录）
- 5 种表格类型智能识别（单列 / 包裹节 / 层级 / KV / 标准）
- 嵌套表格渲染修复：子表格脱离列表缩进，单元格含子表格返回占位文本
- 正文自动转无序列表，图片按内容哈希去重
- 重复运行自动清理输出目录
- OneNote 层级加载稳定检测
- **Flet 桌面 GUI**：暗色/亮色主题、06:00–18:00 自动日间模式、窗口自适应
- **默认输出桌面**：自动读取注册表中用户桌面路径
- **图片语法默认 Obsidian**（`![[]]`）
- **配置记忆**：关闭后保留上次路径和选项

## 快速开始

### 方式一：下载 exe（无需 Python）

从 [Releases](https://github.com/passengerya/OneConvertXmlToMarkdown/releases) 下载 `OneConvert.exe`，双击运行。

### 方式二：源码运行

```bash
pip install flet flet-desktop
双击 Run-OneConvertGUI.bat
```

## 运行环境

- Windows + OneNote（COM 接口）
- 源码运行需 Python 3.11+

## 文件说明

| 文件 | 用途 |
|------|------|
| `OneConvertGUI.py` | Flet 桌面 GUI |
| `Run-OneConvertGUI.bat` | 双击启动 |
| `convert_onenote_xml.py` | XML → Markdown 核心转换器 |
| `Convert-OneNoteSectionToXml.ps1` | .one → 分层 XML（OneNote COM） |
| `Convert-OneNoteToMarkdownPipeline.ps1` | 一键流水线 |
| `release.py` | 交互式打包发布工具 |
| `.oneconvert_config.json` | 用户配置（自动生成） |

## 使用方式

### 图形界面

双击 `Run-OneConvertGUI.bat`。

- 选择 `.one` 文件，输出目录默认桌面 `OneConvertXmlToMarkdown_Output/`
- 桌面路径从注册表自动读取，适配自定义桌面位置
- 图片语法默认 Obsidian，可切换 Markdown
- 所有配置自动保存，重启恢复
- 点击"开始转换"，弹出日志弹窗
- 成功：弹窗 2 秒关闭；失败：弹窗保持打开

### 命令行

```powershell
# 一键流水线
powershell -ExecutionPolicy Bypass -File .\Convert-OneNoteToMarkdownPipeline.ps1 `
  -InputOneFile .\新临检.one `
  -XmlOutputDirectory .\output\xml `
  -MarkdownOutputDirectory .\output\markdown

# 仅 XML
powershell -ExecutionPolicy Bypass -File .\Convert-OneNoteSectionToXml.ps1 `
  -InputOneFile .\新临检.one -OutputDirectory .\output\xml

# 仅 XML → Markdown
python .\convert_onenote_xml.py .\output\xml\新临检 -o .\output\markdown\新临检 --recursive --copy-attachments --asset-dir ../attachment
```

### 打包发布

```bash
python release.py
```

| 选项 | 动作 |
|------|------|
| `[1] 打包` | 仅构建 exe，不改变版本号 |
| `[2] 推送 Release` | 使用已有 exe，tag → push → GitHub Release |
| `[3] 打包并推送` | 完整流水线 |

## 流水线参数

| 参数 | 说明 |
|------|------|
| `-IncludeEmptyPages` | 包含空页面 |
| `-LoadTimeoutSeconds 60` | 加载超时秒数 |
| `-SkipMarkdownStage` | 仅生成 XML |
| `-ImageSyntax markdown\|obsidian` | 图片语法 |
| `-CopyAssets:$false` | 不导出图片 |
| `-AssetDirectoryName attachment` | 资源目录名 |

## 输出结构

```
OneConvertXmlToMarkdown_Output/   # 桌面默认
├── xml/
│   └── 分区名/
│       ├── 001 页面标题/
│       ├── 002 叶子页面.xml
│       └── section-hierarchy.xml
└── markdown/
    ├── 分区名/
    └── attachment/
```

## 注意事项

- 确保 OneNote 能打开目标 `.one` 文件
- 重复导出到同一目录安全
- 大文件建议 `-LoadTimeoutSeconds 120`
- GUI 报错确认 `pip install flet flet-desktop`
- 转换失败弹窗保持打开，可查看错误详情
- exe 由 PyInstaller 打包，杀毒软件可能误报