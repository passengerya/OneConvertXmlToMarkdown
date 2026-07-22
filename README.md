# OneConvertXML

OneNote 分区文件（`.one`）转分层 Markdown，针对嵌套表格深度优化。

`.one → 分层 XML → 分层 Markdown`

## 功能特性

**表格转换**
- 按列数 + 嵌套深度自动路由：1 列 → 有序列表 / 2 列 → 标题层级 / 深层嵌套 → `#` `##` `###` + 列表降维
- 标题规范：一级 `一、` 二级 `（一）` 三级 `1.`

**标注转换**
- 智能识别 OneNote CDATA 和 OE 属性中的 6 种标注：彩色文字、粗体、斜体、删除线、下划线、高亮
- 按颜色值拆分，不同颜色独立映射到 Obsidian 格式
- 粗体 → `**`，斜体 → `*`，删除线 → `~~`，高亮 → `==`，行内代码 → `` ` ``，保留 HTML
- 颜色白→黑自动转换

**桌面 GUI**
- 暗色/亮色主题、06:00–18:00 自动日间、窗口自适应
- 多文件列表式选择，每文件独立移除按钮，浏览追加
- 详细分阶段操作日志，日志弹窗停留 3 秒，支持复制
- 输出目录默认桌面、从注册表读取自定义桌面路径
- 图片语法默认 Obsidian，配置自动记忆

## 快速开始

### 方式一：下载 exe

从 [Releases](https://github.com/passengerya/OneConvertXmlToMarkdown/releases) 下载 `OneConvert.exe`。**无需 Python**，双击运行。

### 方式二：源码运行

```bash
pip install flet flet-desktop
双击 Run-OneConvertGUI.bat
```

## 运行环境

- Windows + OneNote（COM 接口）
- exe 无需 Python；源码运行需 Python 3.8+

## 文件说明

| 文件 | 用途 |
|------|------|
| `OneConvertGUI.py` | Flet 桌面 GUI |
| `Run-OneConvertGUI.bat` | 双击启动 |
| `convert_onenote_xml.py` | XML → Markdown 核心转换器 |
| `Convert-OneNoteSectionToXml.ps1` | .one → XML（OneNote COM） |
| `Convert-OneNoteToMarkdownPipeline.ps1` | 一键流水线 |
| `release.py` | 交互式打包发布（版本号仅推送成功后递增） |
| `.oneconvert_config.json` | 用户配置（自动生成） |

## 使用方式

### 图形界面

双击 `Run-OneConvertGUI.bat`。

- 浏览选择 `.one` / `.xml` 文件，支持多选
- 勾选"标注转换"自动识别标注并弹出映射窗口
- 点击"开始转换"，日志弹窗显示各阶段进度
- 成功弹窗 3 秒关闭；失败弹窗保持，可复制日志调试

### 命令行

```powershell
# 一键流水线
powershell -ExecutionPolicy Bypass -File .\Convert-OneNoteToMarkdownPipeline.ps1 `
  -InputOneFile .\新临检.one -XmlOutputDirectory .\output\xml -MarkdownOutputDirectory .\output\markdown

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

## 输出结构

```
OneConvertXmlToMarkdown_Output/    # 桌面默认
├── xml/
│   └── 分区名/                    # .one 同名 XML 目录
│       ├── 001 页面.xml
│       └── section-hierarchy.xml
└── markdown/
    ├── 分区名/                    # .one → 文件夹（含多个 .md）
    │   └── 001 页面.md
    ├── demo.md                    # .xml → 单个 .md
    └── attachment/                # 共享资源目录
```

## 注意事项

- 确保 OneNote 能打开目标 `.one` 文件
- 重复导出到同一目录安全，自动清理
- 大文件建议 `-LoadTimeoutSeconds 120`
- exe 兼容 Python 3.8–3.14 全版本
- 杀毒软件可能误报，将 exe 加入白名单