#!/usr/bin/env python3
"""OneConvert Pipeline - Flet Desktop GUI (.one -> XML -> Markdown)"""

import json, os, re, shutil, subprocess, sys, threading, winreg
from pathlib import Path
from datetime import datetime

import flet as ft

ROOT = Path(__file__).resolve().parent
# PyInstaller onefile: sys._MEIPASS is the temp extraction dir
if getattr(sys, "frozen", False):
    ROOT = Path(sys._MEIPASS)
PS1  = ROOT / "Convert-OneNoteToMarkdownPipeline.ps1"
ACC  = ft.Colors.BLUE_ACCENT_400


def _desktop_path() -> Path:
    """Read the actual desktop directory from Windows registry."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        )
        val, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        return Path(os.path.expandvars(val))
    except Exception:
        return Path.home() / "Desktop"


DESKTOP = _desktop_path()
OUT_BASE = DESKTOP / "OneConvertXmlToMarkdown_Output"
CONFIG  = ROOT / ".oneconvert_config.json"


def load_config() -> dict:
    try:
        with open(CONFIG, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(cfg: dict):
    try:
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def get_ps() -> str:
    for c in ("pwsh", "powershell"):
        if p := shutil.which(c):
            return p
    for p in (r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
              r"C:\Program Files\PowerShell\7\pwsh.exe"):
        if os.path.isfile(p):
            return p
    return "powershell.exe"


def ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


def main(page: ft.Page):
    page.title = "OneConvert Pipeline"
    page.window.width = 800
    page.window.height = 540
    page.window.min_width = 700
    page.window.min_height = 500
    # Auto theme: light 06:00–18:00, dark otherwise
    hour = datetime.now().hour
    auto_theme = ft.ThemeMode.LIGHT if 6 <= hour < 18 else ft.ThemeMode.DARK
    page.theme_mode = auto_theme
    page.theme = ft.Theme(color_scheme_seed=ft.Colors.INDIGO)
    page.padding = 20

    # -- state ------------------------------------------------------
    proc: subprocess.Popen | None = None
    running = False
    cfg = load_config()

    # -- controls (defaults from saved config or sensible defaults) ---
    txt_input = ft.TextField(
        label="输入文件 (.one / .xml)", hint_text="可多选，以 ; 分隔",
        expand=True, icon=ft.Icons.DESIGN_SERVICES_OUTLINED, dense=True,
        value="",
    )
    txt_xml = ft.TextField(
        label="XML 输出目录",
        value=cfg.get("xml_dir", str(OUT_BASE / "xml")),
        expand=True, icon=ft.Icons.FOLDER_OUTLINED, dense=True,
    )
    txt_md = ft.TextField(
        label="Markdown 输出目录",
        value=cfg.get("md_dir", str(OUT_BASE / "markdown")),
        expand=True, icon=ft.Icons.FOLDER_OUTLINED, dense=True,
    )

    chk_empty  = ft.Checkbox(label="包含空页面", value=cfg.get("chk_empty", False))
    chk_skip   = ft.Checkbox(label="仅生成 XML（跳过 Markdown）", value=cfg.get("chk_skip", False))
    chk_assets = ft.Checkbox(label="复制图片资源", value=cfg.get("chk_assets", True))
    chk_md_only = ft.Checkbox(label="仅生成 Markdown（不保留 XML）", value=cfg.get("chk_md_only", False))
    chk_annotate = ft.Checkbox(label="标注转换（OneNote 标注 → Obsidian）",
                                value=cfg.get("chk_annotate", False))
    ddl_syntax = ft.Dropdown(
        label="图片语法",
        options=[ft.dropdown.Option("markdown"), ft.dropdown.Option("obsidian")],
        value=cfg.get("syntax", "obsidian"), width=140, dense=True,
    )
    txt_asset = ft.TextField(
        label="资源目录名",
        value=cfg.get("asset_dir", "attachment"), width=140, dense=True,
    )

    status_text = ft.Text("就绪", weight=ft.FontWeight.BOLD, size=14)
    progress = ft.ProgressBar(width=160, color=ft.Colors.BLUE_300, visible=False)
    log_lines: list[str] = []

    md_ctrls = [ddl_syntax, chk_assets, txt_asset]
    all_ctrls = [txt_input, txt_xml, txt_md, chk_empty, chk_skip, chk_assets, chk_md_only,
                 chk_annotate, ddl_syntax, txt_asset]

    btn_run = ft.Button(content=ft.Text("开始转换"), icon=ft.Icons.PLAY_ARROW,
                        bgcolor=ACC, color=ft.Colors.WHITE)
    btn_stop = ft.Button(content=ft.Text("停止"), icon=ft.Icons.STOP, visible=False,
                         bgcolor=ft.Colors.TRANSPARENT)

    # -- helpers ----------------------------------------------------
    def save_settings():
        save_config({
            "xml_dir": txt_xml.value or "",
            "md_dir": txt_md.value or "",
            "chk_empty": chk_empty.value,
            "chk_skip": chk_skip.value,
            "chk_assets": chk_assets.value,
            "chk_md_only": chk_md_only.value,
            "chk_annotate": chk_annotate.value,
            "syntax": ddl_syntax.value or "obsidian",
            "asset_dir": txt_asset.value or "attachment",
        })

    def snack(msg: str):
        page.show_dialog(ft.SnackBar(
            ft.Text(msg, color=ft.Colors.ON_ERROR_CONTAINER),
            action="关闭",
            bgcolor=ft.Colors.ERROR_CONTAINER,
            open=True))

    def _push_log(msg: str):
        log_lines.append(msg)

    fp = ft.FilePicker()

    def browse_file(target: ft.TextField, _):
        async def _pick():
            r = await fp.pick_files(allowed_extensions=["one", "xml"], allow_multiple=True)
            if r:
                paths = [f.path for f in r]
                # Append to existing input
                existing = [p.strip() for p in (target.value or "").split(";") if p.strip()]
                for p in paths:
                    if p not in existing:
                        existing.append(p)
                target.value = ";".join(existing)
                page.update()
        page.run_task(_pick)

    def browse_dir(target: ft.TextField, _):
        async def _pick():
            r = await fp.get_directory_path()
            if r:
                target.value = r
                page.update()
        page.run_task(_pick)

    def open_dir(target: ft.TextField, _):
        p = (target.value or "").strip()
        if p and os.path.isdir(p) and sys.platform == "win32":
            os.startfile(p)
        else:
            snack("目录不存在")

    def on_skip(_):
        for c in md_ctrls:
            c.disabled = chk_skip.value
        page.update()

    def toggle_theme(_):
        page.theme_mode = ft.ThemeMode.LIGHT if page.theme_mode == ft.ThemeMode.DARK else ft.ThemeMode.DARK
        page.update()

    def show_about(_):
        page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Icon(ft.Icons.ART_TRACK, color=ACC),
                ft.Text("OneConvert Pipeline"),
            ]),
            content=ft.Column([
                ft.Text(".one -> XML -> Markdown 转换工具", size=14),
                ft.Divider(),
                ft.Text("Flet 桌面应用 · 嵌套表格修复版", size=12, color=ft.Colors.GREEN_400),
                ft.Text("依赖: OneNote (COM) + Python 3", size=12, color=ft.Colors.SECONDARY),
            ], spacing=8, tight=True),
            actions=[ft.Button(content=ft.Text("确定"), on_click=lambda e: page.pop_dialog(),
                               bgcolor=ACC, color=ft.Colors.WHITE)],
        ))

    def show_log_dialog(is_error: bool = False):
        """Show the conversion log in a dialog. Auto-close after 2s unless error."""
        log_content = ft.Column([], scroll=ft.ScrollMode.AUTO, expand=True)
        for msg in log_lines[-200:]:
            c = ft.Colors.ERROR if "[ERR]" in msg else None
            log_content.controls.append(
                ft.Text(msg, size=12, font_family="Consolas", selectable=True, color=c))

        def _close():
            try:
                page.pop_dialog()
            except Exception:
                pass

        def _copy_log(e):
            page.set_clipboard("\n".join(log_lines))
            page.pop_dialog()
            page.update()

        footer = ("转换失败，请查看上方错误信息" if is_error
                  else "此窗口将在 2 秒后自动关闭")
        actions = [ft.Text(footer, size=12, color=ft.Colors.ERROR if is_error else ft.Colors.SECONDARY)]
        if is_error:
            actions.insert(0, ft.Button(content=ft.Text("复制日志"), icon=ft.Icons.COPY,
                                          on_click=_copy_log, bgcolor=ACC, color=ft.Colors.WHITE))
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Icon(ft.Icons.TERMINAL_OUTLINED, color=ft.Colors.ERROR if is_error else ACC),
                ft.Text("运行日志" if not is_error else "运行日志 — 转换失败"),
            ], spacing=8),
            content=ft.Container(
                content=log_content,
                width=700, height=400,
                bgcolor=ft.Colors.BLACK26 if page.theme_mode == ft.ThemeMode.DARK else ft.Colors.BLACK12,
                border=ft.Border.all(1, ft.Colors.OUTLINE),
                border_radius=8,
                padding=12,
            ),
            actions=actions,
            actions_alignment=ft.MainAxisAlignment.CENTER,
        )
        page.show_dialog(dialog)

        if not is_error:
            def auto_close():
                import time
                time.sleep(2)
                try:
                    page.run_thread(_close)
                except Exception:
                    pass
            threading.Thread(target=auto_close, daemon=True).start()

    # -- annotation conversion: OneNote formatting → Obsidian ----
    BINARY_TYPES = [  # on/off types — no color value
        ("粗体",   r'font-weight:bold',              r'<b>', r'</b>'),
        ("斜体",   r'font-style:italic',             r'<i>', r'</i>'),
        ("删除线", r'text-decoration:line-through',  r'<s>', r'</s>'),
        ("下划线", r'text-decoration:underline',     r'<u>', r'</u>'),
    ]
    COLOR_TYPES = [  # per-color types — scan extracts color values
        ("彩色文字", r'color:(#[a-fA-F0-9]{3,6})',         r'<font color="\1">', r'</font>'),
        ("彩色高亮", r'background(?:-color)?:(\S+?)[;"]',  r'<span style="background-color:\1">', r'</span>'),
    ]
    OBSIDIAN_TARGETS = [
        ("不转换", ""),
        ("**粗体**", "**"),
        ("*斜体*", "*"),
        ("~~删除线~~", "~~"),
        ("==高亮==", "=="),
        ("`行内代码`", "`"),
        ("<u>下划线</u>", "<u>"),
        ("<font>彩色</font>", "<font>"),
    ]

    def _scan_items(xml_dirs: list[Path]) -> list[tuple[str, str, str, str]]:
        """Scan XML for annotation items. Returns [(label, color_value, open_pat, close_pat), ...]."""
        items: list[tuple[str, str, str, str]] = []
        seen: set[str] = set()
        for xml_src in xml_dirs:
            try:
                fps = [xml_src] if xml_src.is_file() else list(xml_src.rglob("*.xml"))
                for xf in fps:
                    t = xf.read_text(encoding="utf-8", errors="replace")
                    # Binary types
                    for label, scan_pat, open_pat, close_pat in BINARY_TYPES:
                        if re.search(scan_pat, t) and label not in seen:
                            seen.add(label)
                            items.append((label, "", open_pat, close_pat))
                    # Per-color types
                    for label, scan_pat, open_tmpl, close_pat in COLOR_TYPES:
                        for m in re.finditer(scan_pat, t):
                            val = m.group(1).strip().rstrip(";").strip('"').strip("'")
                            key = f"{label}({val})"
                            if key not in seen:
                                seen.add(key)
                                open_pat = open_tmpl.replace(r'\1', re.escape(val))
                                items.append((key, val, open_pat, close_pat))
            except Exception:
                pass
        return items

    def _show_annotation_dialog(callback, items: list):
        """Show per-item annotation mapping dialog."""
        if not items:
            snack("未检测到任何标注类型（彩色、粗体、斜体、删除线、下划线、高亮）")
            callback()
            return
        dd_map: dict[str, ft.Dropdown] = {}
        rows = []
        for label, val, open_pat, close_pat in items:
            display = f"{label}  " if val else label
            dd = ft.Dropdown(
                options=[ft.dropdown.Option(n) for n, _ in OBSIDIAN_TARGETS],
                value="不转换", width=160, dense=True,
            )
            dd_map[label] = dd
            color_hint = ft.Container(width=16, height=16, border_radius=8, bgcolor=val) if val else ft.Text("")
            rows.append(ft.Row([
                color_hint,
                ft.Text(display, size=12, width=160),
                ft.Text("→", size=12),
                dd,
            ], spacing=4))

        def _confirm(e):
            result = {}
            for label, dd in dd_map.items():
                sel = dd.value
                if not sel or sel == "不转换":
                    continue
                for t_name, t_marker in OBSIDIAN_TARGETS:
                    if t_name == sel and t_marker:
                        result[label] = t_marker
                        break
            page.pop_dialog()
            page.update()
            callback(result)

        amap: dict[str, str] = {}

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text(f"标注转换设置 ({len(items)} 项)"),
            content=ft.Column([
                ft.Text("检测到以下标注，选择目标格式：", size=13),
                ft.Divider(),
            ] + rows, spacing=4, scroll=ft.ScrollMode.AUTO),
            actions=[
                ft.Button(content=ft.Text("跳过"), on_click=lambda e: (page.pop_dialog(), callback())),
                ft.Button(content=ft.Text("确认"), on_click=_confirm,
                          bgcolor=ACC, color=ft.Colors.WHITE),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.show_dialog(dlg)
        page.update()

    def _apply_annotations(md_text: str, amap: dict, items: list) -> str:
        """Apply per-item annotation mappings."""
        for label, val, open_pat, close_pat in items:
            marker = amap.get(label)
            if not marker or marker == "<font>":
                continue
            md_text = re.sub(open_pat + r'(.*?)' + close_pat,
                             marker + r'\1' + marker,
                             md_text, flags=re.DOTALL)
        return md_text

    # -- pipeline (multi-file: PS for each .one, in-process for Markdown) --
    def run(_):
        nonlocal proc, running
        if running:
            return

        in_raw = (txt_input.value or "").strip()
        xout_base = (txt_xml.value or "").strip() or str(OUT_BASE / "xml")
        mout_base = (txt_md.value or "").strip() or str(OUT_BASE / "markdown")
        adir = (txt_asset.value or "").strip()

        if not in_raw:           snack("请选择 .one 或 .xml 文件"); return
        if not chk_assets.value and not adir: snack("资源目录名不能为空"); return

        files = [p.strip() for p in in_raw.split(";") if p.strip()]
        ones = [f for f in files if f.lower().endswith(".one")]
        xmls = [f for f in files if f.lower().endswith(".xml")]
        if not ones and not xmls: snack("必须是 .one 或 .xml 文件"); return
        for f in files:
            if not os.path.isfile(f):
                snack(f"未找到: {f}"); return

        def _start():
            nonlocal proc, running
            running = True
            log_lines.clear()
            status_text.value = "处理中..."
            status_text.color = ft.Colors.ORANGE_400
            progress.value = None
            progress.visible = True
            for c in all_ctrls:
                c.disabled = True
            btn_run.disabled = True
            btn_stop.visible = True
            page.update()

            log_lines.append(f"[{ts()}] 开始转换")
            log_lines.append(f"文件数: {len(files)} ({len(ones)} one, {len(xmls)} xml)")
            if not chk_skip.value:
                log_lines.append(f"MD:   {mout_base}")
                log_lines.append(f"语法: {ddl_syntax.value}")
            log_lines.append("-" * 50)

            xml_script = str(ROOT / "Convert-OneNoteSectionToXml.ps1")
            syntax = ddl_syntax.value or "obsidian"
            asset_rel = f"../{adir}" if adir and "/" not in adir and "\\" not in adir else adir

            def worker():
                nonlocal proc
                code = 0; total_md = 0; ext_imgs = 0; pl_imgs = 0
                all_written: list[Path] = []
                xml_dirs: list[Path] = []
                annot_items: list = []
                try:
                    sys.path.insert(0, str(ROOT))
                    import convert_onenote_xml as converter

                    # ── Phase 1: .one → XML ──
                    for one_path in ones:
                        stem = Path(one_path).stem
                        log_lines.append(f"[{ts()}] .one -> XML: {stem}")
                        ps_args = [
                            get_ps(), "-NoProfile", "-ExecutionPolicy", "Bypass",
                            "-File", xml_script,
                            "-InputOneFile", one_path,
                            "-OutputDirectory", str(Path(xout_base)),
                            "-LoadTimeoutSeconds", "30",
                        ]
                        if chk_empty.value:
                            ps_args.append("-IncludeEmptyPages")
                        p1 = subprocess.Popen(
                            ps_args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, encoding="utf-8", errors="replace", cwd=str(ROOT),
                            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
                        )
                        proc = p1
                        out1, err1 = p1.communicate(timeout=180)
                        if out1:
                            for line in out1.split("\n"):
                                if line.strip(): page.run_thread(_push_log, line.strip())
                        if err1:
                            for line in err1.split("\n"):
                                if line.strip(): page.run_thread(_push_log, f"[ERR] {line.strip()}")
                        if p1.returncode != 0:
                            log_lines.append(f"[ERR] {stem} XML 转换失败")
                            code = max(code, p1.returncode)
                            continue
                        xml_dir = Path(xout_base) / stem
                        if xml_dir.exists():
                            xml_dirs.append(xml_dir)
                            log_lines.append(f"  {stem} XML done")
                        else:
                            log_lines.append(f"[ERR] XML 目录不存在: {xml_dir}")

                    # XML input files are their own source
                    for xml_path in xmls:
                        xml_dirs.append(Path(xml_path))

                    # ── Phase 2: scan XML for annotation types ──
                    if chk_annotate.value and xml_dirs:
                        annot_items = _scan_items(xml_dirs)
                        msg = f"检测到 {len(annot_items)} 种标注" if annot_items else "未检测到标注类型"
                        log_lines.append(f"  {msg}")
                        status_text.value = msg
                        page.update()

                    # ── Phase 3: XML → MD ──
                    if not chk_skip.value:
                        for xml_src in xml_dirs:
                            if xml_src.is_file():
                                # Single XML file
                                mout_all = Path(mout_base)
                                mout_all.mkdir(parents=True, exist_ok=True)
                                written, resolver = converter.convert_path(
                                    xml_input=xml_src, output_dir=mout_all,
                                    attachment_dir=None, sort_key="name",
                                    recursive=False, image_syntax=syntax,
                                    copy_attachments=chk_assets.value,
                                    asset_dir_name=asset_rel,
                                )
                            else:
                                # XML directory (from .one)
                                stem = xml_src.name
                                mout_one = Path(mout_base) / stem
                                mout_one.mkdir(parents=True, exist_ok=True)
                                written, resolver = converter.convert_path(
                                    xml_input=xml_src, output_dir=mout_one,
                                    attachment_dir=None, sort_key="name",
                                    recursive=True, image_syntax=syntax,
                                    copy_attachments=chk_assets.value,
                                    asset_dir_name=asset_rel,
                                )
                            total_md += len(written)
                            ext_imgs += resolver.extracted_from_xml_count
                            pl_imgs += resolver.placeholder_count
                            log_lines.append(f"  {xml_src.name} -> {len(written)} md")
                            all_written.extend(written)
                            if chk_md_only.value and xml_src.is_dir():
                                try: shutil.rmtree(xml_src)
                                except Exception: pass

                    if chk_md_only.value and chk_skip.value:
                        try:
                            shutil.rmtree(Path(xout_base))
                            log_lines.append("  XML 已清理")
                        except Exception: pass
                    if not chk_skip.value:
                        log_lines.append(f"  总计: {total_md} md 文件")
                        if ext_imgs: log_lines.append(f"  提取图片: {ext_imgs}")
                        if pl_imgs: log_lines.append(f"  占位图: {pl_imgs}")
                except subprocess.TimeoutExpired:
                    log_lines.append("[ERR] 转换超时(180s)")
                    code = 1
                except Exception as ex:
                    log_lines.append(f"[ERR] {ex}")
                    code = 1

                # ── Phase 4: apply annotations ──
                if chk_annotate.value:
                    def _apply_and_finish(amap=None):
                        if amap:
                            for md_path in all_written:
                                try:
                                    t = md_path.read_text(encoding="utf-8-sig")
                                    md_path.write_text(_apply_annotations(t, amap, annot_items), encoding="utf-8-sig")
                                except Exception:
                                    pass
                        show_log_dialog()
                        page.run_thread(finish, code)
                    page.run_thread(lambda: _show_annotation_dialog(_apply_and_finish, annot_items))
                else:
                    page.run_thread(lambda: (show_log_dialog(), finish(code)))

            def finish(c: int):
                nonlocal running
                running = False
                btn_run.disabled = False
                btn_stop.visible = False
                if c == 0:
                    status_text.value = "已完成"
                    status_text.color = ft.Colors.GREEN_400
                    progress.value = 1.0; progress.color = ft.Colors.GREEN_300
                else:
                    status_text.value = f"失败 (code={c})"
                    status_text.color = ft.Colors.ERROR
                    progress.value = 0; progress.color = ft.Colors.ERROR
                    log_lines.append("-" * 50)
                    log_lines.append(f"[{ts()}] 结束, 退出码: {c}")
                    show_log_dialog(is_error=True)
                for ctrl in all_ctrls: ctrl.disabled = False
                on_skip(None)
                page.update()

            threading.Thread(target=worker, daemon=True).start()

        _start()

    def stop(_):
        nonlocal proc
        if proc and not proc.poll():
            proc.kill()
            status_text.value = "已终止"
            status_text.color = ft.Colors.ERROR
            page.update()

    # -- wire events -----------------------------------------------
    chk_skip.on_change = on_skip
    btn_run.on_click = run
    btn_stop.on_click = stop

    # Save settings on change
    for ctrl in [txt_input, txt_xml, txt_md, chk_empty, chk_skip, chk_assets, chk_md_only, chk_annotate, txt_asset]:
        orig = ctrl.on_change
        def _wrap(c, o):
            c.on_change = lambda e: (save_settings(), o(e) if o else None)
        _wrap(ctrl, orig)
    ddl_syntax.on_select = lambda e: save_settings()

    # ==============================================================
    #  LAYOUT — single page, no scroll, no inline log
    # ==============================================================

    page.add(
        ft.Column([
            ft.Container(content=ft.Column([
                # Header — icon + title + theme & about buttons
                ft.Row([
                    ft.Icon(ft.Icons.ART_TRACK, size=36, color=ACC),
                    ft.Column([
                        ft.Text("OneConvert Pipeline", weight=ft.FontWeight.BOLD, size=20),
                        ft.Text(".one -> XML -> Markdown   |   嵌套表格转换", size=12,
                                color=ft.Colors.SECONDARY),
                    ], spacing=2),
                    ft.Container(expand=True),
                    ft.IconButton(icon=ft.Icons.BRIGHTNESS_4_OUTLINED, tooltip="切换主题", on_click=toggle_theme),
                    ft.IconButton(icon=ft.Icons.INFO_OUTLINED, tooltip="关于", on_click=show_about),
                ]),
                ft.Divider(height=16),

                # File paths
                ft.Text("文件路径", weight=ft.FontWeight.BOLD, size=14),
                ft.Row([txt_input, ft.Button(content=ft.Text("浏览"), icon=ft.Icons.SEARCH,
                                              on_click=lambda _: browse_file(txt_input, _))]),
                ft.Container(height=6),
                ft.Row([txt_xml,  ft.Button(content=ft.Text("浏览"), icon=ft.Icons.SEARCH,
                                              on_click=lambda _: browse_dir(txt_xml, _)),
                        ft.IconButton(icon=ft.Icons.FOLDER_OPEN, tooltip="打开目录",
                                      on_click=lambda _: open_dir(txt_xml, _), icon_color=ACC)]),
                ft.Container(height=6),
                ft.Row([txt_md,   ft.Button(content=ft.Text("浏览"), icon=ft.Icons.SEARCH,
                                              on_click=lambda _: browse_dir(txt_md, _)),
                        ft.IconButton(icon=ft.Icons.FOLDER_OPEN, tooltip="打开目录",
                                      on_click=lambda _: open_dir(txt_md, _), icon_color=ACC)]),
                ft.Divider(height=16),

                # Options
                ft.Text("转换选项", weight=ft.FontWeight.BOLD, size=14),
                ft.Container(height=8),
                ft.Card(content=ft.Container(content=ft.Column([
                    ft.Row([chk_empty, chk_assets, ft.Container(expand=True)],
                           alignment=ft.MainAxisAlignment.START),
                    chk_skip,
                    chk_md_only,
                    chk_annotate,
                    ft.Divider(height=8),
                    ft.Row([ddl_syntax, txt_asset, ft.Container(expand=True)],
                           alignment=ft.MainAxisAlignment.START, spacing=16),
                ], spacing=8), padding=20)),
                ft.Divider(height=16),

                # Actions + status
                ft.Row([
                    btn_run,
                    ft.Container(expand=True),
                    btn_stop,
                ], spacing=12, alignment=ft.MainAxisAlignment.START),
                ft.Container(height=8),
                ft.Row([
                    ft.Row([
                        ft.Icon(ft.Icons.CIRCLE, size=10, color=ft.Colors.GREEN_400),
                        status_text,
                    ], spacing=6),
                    progress,
                ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            ], spacing=8), padding=20, expand=True),
        ], expand=True, scroll=ft.ScrollMode.AUTO, alignment=ft.MainAxisAlignment.START),
    )

    # Auto-load default .one (only first run)
    if not txt_input.value:
        default = ROOT / "新临检.one"
        if default.exists():
            txt_input.value = str(default)


if __name__ == "__main__":
    ft.run(main, view=ft.AppView.FLET_APP)