# OneConvertXML 使用说明

## 项目用途

本项目用于将 OneNote 导出的分区文件（`.one`）按页面层级转换为：
（主要用于OneNote嵌套表格的转换）
1. 分层 XML（章节文件夹 + 页面 XML）
2. 分层 Markdown（保持层级结构）

推荐使用一键流水线：

`.one -> 分层 XML -> 分层 Markdown`

## 当前功能特性（最终版）

- `.one` 转分层 XML（按页面层级建目录）
- XML 再转分层 Markdown
- 自动跳过 `section-hierarchy.xml`（不会生成 `section-hierarchy.md`）
- 图片按内容去重（同一图片不会重复导出多份）
- 图片默认导出到 `attachment` 文件夹（与 Markdown 大文件夹同级）
- 正文自动转无序列表（`- `）
- 清理 Markdown 无效空行（列表/表格/标题后的多余空行）
- 重复运行时自动清理该分区输出目录，避免生成 ` (2)` 副本
- OneNote 层级加载采用“稳定检测”，降低漏页/残缺导出风险

## 运行环境

- Windows（需安装 OneNote，支持 COM 接口）
- Python 3（用于 XML -> Markdown）
- PowerShell（Windows 自带）

## 文件说明

- `Convert-OneNoteSectionToXml.ps1`
  - 第一步：`.one -> 分层 XML`
- `convert_onenote_xml.py`
  - 第二步：`分层 XML -> 分层 Markdown`
- `Convert-OneNoteToMarkdownPipeline.ps1`
  - 一键流水线脚本（串联两步）
- `OneConvertPipeline-GUI.ps1`
  - 图形界面（WinForms）
- `Run-OneConvertPipeline-GUI.bat`
  - 双击启动图形界面

## 默认保存位置（已内置）

如果不手动指定输出目录，默认保存到项目根目录下：

- `output\xml`：存放 XML 相关文件
- `output\markdown`：存放 Markdown 相关文件

如果这些目录不存在，会自动创建。

如果你在 GUI 或命令行中自定义了输出目录，则保存到你指定的位置。

## 推荐使用方式（图形界面）

双击运行：

`Run-OneConvertPipeline-GUI.bat`

界面支持：

- 选择输入 `.one` 文件
- 设置 XML 输出目录（默认 `output\xml`）
- 设置 Markdown 输出目录（默认 `output\markdown`）
- 配置超时、空页、图片语法、资源导出等参数
- 实时查看日志

## 命令行用法

### 1) 一键流水线（推荐）

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

### 3) 单独执行 XML -> Markdown

```powershell
python .\convert_onenote_xml.py .\output\xml\新临检 -o .\output\markdown\新临检 --recursive --copy-attachments --asset-dir ../attachment
```

说明：
- 流水线脚本会自动把资源目录设置为与 Markdown 大文件夹同级的 `attachment`
- 直接调用 `convert_onenote_xml.py` 时，若想得到同样结构，请显式指定 `--asset-dir ../attachment`

## 常用参数（流水线脚本）

- `-IncludeEmptyPages`
  - XML 阶段包含空页面（默认跳过空页面）
- `-LoadTimeoutSeconds 60`
  - 等待 OneNote 加载完整层级的超时秒数（大文件建议调大）
- `-SkipMarkdownStage`
  - 只执行 `.one -> XML`
- `-ImageSyntax markdown|obsidian`
  - Markdown 图片语法
- `-CopyAssets:$false`
  - 不导出图片资源目录（默认导出）
- `-AssetDirectoryName attachment`
  - 资源目录名称（流水线会自动放到 Markdown 大文件夹同级）

## 输出结构说明（默认）

### XML 输出（`output\xml`）

- `output\xml\新临检\...`
- 有子页面的页面（章节）创建文件夹
- 叶子页面（节）直接输出为 `页面标题.xml`
- 额外输出 `section-hierarchy.xml`（分区层级元数据）

### Markdown 输出（`output\markdown`）

- `output\markdown\新临检\...`（Markdown 文件）
- `output\markdown\attachment\...`（图片资源）
- Markdown 与 XML 层级结构一致
- 自动跳过 `section-hierarchy.xml`

## 注意事项

- 请确保 OneNote 能正常打开目标 `.one` 文件
- 中文 GUI 建议优先通过 `.bat` 启动器启动
- 重复导出到同一目录是安全的：程序会自动清理该分区输出目录后重新生成
- 若仍怀疑漏页或残缺内容，请增大 `-LoadTimeoutSeconds`（例如 `60` / `120`）
