# OneConvertXML

将 OneNote 分区文件（`.one`）按页面层级转换为分层 Markdown，针对嵌套表格深度优化。

`.one → 分层 XML → 分层 Markdown`

## 功能特性

- 按列数 + 嵌套深度自动选渲染策略（单列 → 有序列表 / 双列 → 标题层级 / 深层嵌套 → `#` → `##` → `###` → 列表降维）
- 标题规范：一级 `一、` / 二级 `（一）` / 三级 `1.`
- 图片独立段落提取、颜色白→黑转换、OneNote 标记原样保留
- Flet 桌面 GUI：暗色/亮色主题、06:00–18:00 自动日间、窗口自适应
- 默认输出桌面 `OneConvertXmlToMarkdown_Output/`，路径从注册表自动读取
- 图片语法默认 Obsidian（`![[]]`）
- 配置自动记忆，重启恢复
- 错误弹窗常驻 + 复制日志按钮，日间/夜间模式文字可读

## 快速开始

### 方式一：下载 exe

从 [Releases](https://github.com/passengerya/OneConvertXmlToMarkdown/releases) 下载 `OneConvert.exe`，双击运行。**无需安装 Python**，仅需 Windows + OneNote。

### 方式二：源码运行

```bash
pip install flet flet-desktop
双击 Run-OneConvertGUI.bat
```

## 运行环境

- Windows + OneNote（COM 接口）
- Python 3.8+（exe 无需 Python）

## 文件说明

| 文件 | 用途 |
|------|------|
| `OneConvertGUI.py` | Flet 桌面 GUI |
| `Run-OneConvertGUI.bat` | 双击启动 |
| `convert_onenote_xml.py` | XML → Markdown 核心转换器 |
| `Convert-OneNoteSectionToXml.ps1` | .one → 分层 XML（OneNote COM） |
| `Convert-OneNoteToMarkdownPipeline.ps1` | 一键流水线 |
| `release.py` | 交互式打包发布工具（版本号仅推送成功后递增） |
| `.oneconvert_config.json` | 用户配置（自动生成） |

## 使用方式

### 图形界面

双击 `Run-OneConvertGUI.bat`。

- 选择 `.one` 文件，输出目录默认桌面
- 图片语法默认 Obsidian，可切换 Markdown
- 点击"开始转换"，弹出日志弹窗
- 成功：弹窗 2 秒关闭；失败：弹窗保持打开，可复制日志

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

| 选项 | 动作 | 版本号 |
|------|------|--------|
| `[1] 打包` | 仅构建 exe | 不变 |
| `[2] 推送 Release` | 推送成功 → tag → GitHub Release | +1 |
| `[3] 打包并推送` | 打包 + 推送 | +1 |

## 嵌套表格处理规则

| 表格类型 | 左列 | 右列 |
|----------|------|------|
| 1 列 | — | 有序列表 |
| 2 列无嵌套 | `# 一、标题` | 原文 |
| 2 列 1 层嵌套 | `# 一、标题` | 保留表格 |
| 3 层+ | `#` → `##` → `###` 逐级降 | 正文或列表 |
| 4 层+ | 无序/有序列表 | 按缩进保持层级 |

## 注意事项

- 确保 OneNote 能打开目标 `.one` 文件
- 重复导出到同一目录安全，会自动清理
- 大文件建议 `-LoadTimeoutSeconds 120`
- exe 由 PyInstaller 打包，兼容 Python 3.8–3.14 全版本
- 杀毒软件可能误报，将 exe 加入白名单