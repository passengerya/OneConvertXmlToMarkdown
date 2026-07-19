# OneConvertXML 使用说明

## 项目用途

将 OneNote 分区文件（`.one`）按页面层级转换为分层 Markdown，针对嵌套表格深度优化。

`.one → 分层 XML → 分层 Markdown`

## 当前功能特性

- `.one` 转分层 XML（按页面层级建目录）
- XML 转分层 Markdown，5 种表格类型智能识别（单列 / 包裹节 / 层级 / KV / 标准）
- 嵌套表格深度渲染修复：子表格脱离列表缩进独立显示，单元格含子表格返回占位文本
- 正文自动转无序列表
- 图片按内容哈希去重，从 XML 嵌入式二进制数据提取
- 重复运行时自动清理输出目录
- OneNote 层级加载采用稳定检测（areAllPagesAvailable + 页数稳定确认）
- Flet 桌面 GUI，暗色主题，日间/夜间切换，窗口自适应

## 运行环境

- Windows（需安装 OneNote，支持 COM 接口）
- Python 3.11+
- PowerShell（Windows 自带）

```bash
pip install flet flet-desktop
```

## 文件说明

| 文件 | 用途 |
|------|------|
| `OneConvertGUI.py` | Flet 桌面 GUI |
| `Run-OneConvertGUI.bat` | 双击启动 |
| `convert_onenote_xml.py` | XML → Markdown 核心转换器 |
| `Convert-OneNoteSectionToXml.ps1` | .one → 分层 XML（OneNote COM） |
| `Convert-OneNoteToMarkdownPipeline.ps1` | 一键流水线（串联两步） |

## 使用方式 — 图形界面（推荐）

双击 `Run-OneConvertGUI.bat`。

- 选择 `.one` 文件，设置 XML / Markdown 输出目录
- 配置空页面、图片语法、资源导出参数
- 点击"开始转换"，弹出运行日志弹窗
- 正常完成：弹窗 2 秒后自动关闭
- 转换失败：弹窗保持打开，显示错误详情
- 日间/夜间模式切换、关于信息

## 命令行用法

### 一键流水线

```powershell
powershell -ExecutionPolicy Bypass -File .\Convert-OneNoteToMarkdownPipeline.ps1 `
  -InputOneFile .\新临检.one `
  -XmlOutputDirectory .\output\xml `
  -MarkdownOutputDirectory .\output\markdown
```

### 仅生成分层 XML

```powershell
powershell -ExecutionPolicy Bypass -File .\Convert-OneNoteSectionToXml.ps1 `
  -InputOneFile .\新临检.one `
  -OutputDirectory .\output\xml
```

### 单独执行 XML → Markdown

```powershell
python .\convert_onenote_xml.py .\output\xml\新临检 -o .\output\markdown\新临检 --recursive --copy-attachments --asset-dir ../attachment
```

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
output/
├── xml/
│   └── 分区名/
│       ├── 001 页面标题/      # 含子页面 → 文件夹
│       ├── 002 叶子页面.xml   # 叶子页面 → XML
│       └── section-hierarchy.xml
└── markdown/
    ├── 分区名/                # Markdown（同层级）
    └── attachment/            # 图片资源
```

## 注意事项

- 确保 OneNote 能正常打开目标 `.one` 文件
- 重复导出到同一目录安全，会自动清理
- 大文件建议 `-LoadTimeoutSeconds 120`
- GUI 报错请确认 `pip install flet flet-desktop`
- 转换失败时会弹出错误详情窗口，不会自动关闭