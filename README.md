# OneConvertXML 使用说明

## 项目用途

将 OneNote 分区文件（`.one`）按页面层级转换为：
1. 分层 XML（章节文件夹 + 页面 XML）
2. 分层 Markdown（保持层级结构，针对嵌套表格深度优化）

`.one → 分层 XML → 分层 Markdown`

## 当前功能特性

- `.one` 转分层 XML（按页面层级建目录）
- XML 再转分层 Markdown
- 嵌套表格深度渲染修复：子表格在列表环境中脱离缩进独立显示，单元格含子表格时返回占位文本，避免 Markdown 渲染错乱
- 自动跳过 `section-hierarchy.xml`（不会生成 `section-hierarchy.md`）
- 图片按内容去重（同一图片不会重复导出多份）
- 图片默认导出到 `attachment` 文件夹（与 Markdown 大文件夹同级）
- 正文自动转无序列表（`- `）
- 清理 Markdown 无效空行（列表/表格/标题后的多余空行）
- 重复运行时自动清理该分区输出目录，避免生成 ` (2)` 副本
- OneNote 层级加载采用"稳定检测"，降低漏页/残缺导出风险

## 运行环境

- Windows（需安装 OneNote，支持 COM 接口）
- Python 3.11+（含 `flet` 包）
- PowerShell（Windows 自带）

首次使用需安装 Flet：

```bash
pip install flet flet-desktop
```

## 文件说明

| 文件 | 用途 |
|------|------|
| `OneConvertGUI.py` | Flet 桌面图形界面（现代化 UI，暗色主题，自适应窗口） |
| `Run-OneConvertGUI.bat` | 双击启动图形界面 |
| `Convert-OneNoteSectionToXml.ps1` | 第一步：`.one → 分层 XML` |
| `convert_onenote_xml.py` | 第二步：`分层 XML → 分层 Markdown` |
| `Convert-OneNoteToMarkdownPipeline.ps1` | 一键流水线脚本（串联两步） |

## 推荐使用方式 — 图形界面

双击 `Run-OneConvertGUI.bat` 启动 Flet 桌面应用。

界面功能：
- 选择输入 `.one` 文件
- 设置 XML / Markdown 输出目录
- 配置空页、图片语法、资源导出等参数
- 点击"开始转换"弹出实时日志
- 暗色/亮色主题切换
- 窗口自适应，小窗口可滚动

## 命令行用法

### 1) 一键流水线

```powershell
powershell -ExecutionPolicy Bypass -File .\Convert-OneNoteToMarkdownPipeline.ps1 `
  -InputOneFile .\新临检.one `
  -XmlOutputDirectory .\output\xml `
  -MarkdownOutputDirectory .\output\markdown
```

### 2) 仅生成分层 XML

```powershell
powershell -ExecutionPolicy Bypass -File .\Convert-OneNoteSectionToXml.ps1 `
  -InputOneFile .\新临检.one `
  -OutputDirectory .\output\xml
```

### 3) 单独执行 XML → Markdown

```powershell
python .\convert_onenote_xml.py .\output\xml\新临检 -o .\output\markdown\新临检 --recursive --copy-attachments --asset-dir ../attachment
```

说明：
- 流水线脚本自动将资源目录设为与 Markdown 大文件夹同级的 `attachment`
- 直接调用 `convert_onenote_xml.py` 时，若需同样结构，请显式指定 `--asset-dir ../attachment`

## 常用参数（流水线脚本）

| 参数 | 说明 |
|------|------|
| `-IncludeEmptyPages` | XML 阶段包含空页面（默认跳过） |
| `-LoadTimeoutSeconds 60` | OneNote 加载超时秒数（大文件建议调大） |
| `-SkipMarkdownStage` | 仅执行 `.one → XML` |
| `-ImageSyntax markdown\|obsidian` | Markdown 图片语法 |
| `-CopyAssets:$false` | 不导出图片资源目录（默认导出） |
| `-AssetDirectoryName attachment` | 资源目录名称 |

## 输出结构说明（默认）

```
output/
├── xml/
│   └── 分区名/
│       ├── 001 页面标题/     # 含子页面的章节 → 文件夹
│       ├── 002 叶子页面.xml  # 叶子页面 → XML 文件
│       └── section-hierarchy.xml
└── markdown/
    ├── 分区名/               # Markdown 文件（与 XML 同层级）
    └── attachment/           # 图片资源
```

## 注意事项

- 请确保 OneNote 能正常打开目标 `.one` 文件
- 重复导出到同一目录是安全的：程序会自动清理该分区输出目录后重新生成
- 若怀疑漏页或残缺内容，请增大 `-LoadTimeoutSeconds`（例如 `60` / `120`）
- GUI 启动报错请确认已安装 `flet` 和 `flet-desktop`：`pip install flet flet-desktop`