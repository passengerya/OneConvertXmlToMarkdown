#!/usr/bin/env python3
"""OneConvert Pipeline - Flet Desktop GUI (.one -> XML -> Markdown)"""

import os, shutil, subprocess, sys, threading
from pathlib import Path
from datetime import datetime

import flet as ft

ROOT = Path(__file__).resolve().parent
PS1  = ROOT / "Convert-OneNoteToMarkdownPipeline.ps1"
ACC  = ft.Colors.BLUE_ACCENT_400


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
    page.theme_mode = ft.ThemeMode.DARK
    page.theme = ft.Theme(color_scheme_seed=ft.Colors.INDIGO)
    page.padding = 20

    # -- state ------------------------------------------------------
    proc: subprocess.Popen | None = None
    running = False

    # -- controls ---------------------------------------------------
    txt_input = ft.TextField(
        label="输入 .one 文件", hint_text="点击浏览选择 OneNote 分区文件",
        expand=True, icon=ft.Icons.DESIGN_SERVICES_OUTLINED, dense=True,
    )
    txt_xml = ft.TextField(
        label="XML 输出目录", value=str(ROOT / "output" / "xml"),
        expand=True, icon=ft.Icons.FOLDER_OUTLINED, dense=True,
    )
    txt_md = ft.TextField(
        label="Markdown 输出目录", value=str(ROOT / "output" / "markdown"),
        expand=True, icon=ft.Icons.FOLDER_OUTLINED, dense=True,
    )

    chk_empty  = ft.Checkbox(label="包含空页面", value=False)
    chk_skip   = ft.Checkbox(label="仅生成 XML（跳过 Markdown）", value=False)
    chk_assets = ft.Checkbox(label="复制图片资源", value=True)
    ddl_syntax = ft.Dropdown(
        label="图片语法",
        options=[ft.dropdown.Option("markdown"), ft.dropdown.Option("obsidian")],
        value="markdown", width=140, dense=True,
    )
    txt_asset = ft.TextField(label="资源目录名", value="attachment", width=140, dense=True)

    status_text = ft.Text("就绪", weight=ft.FontWeight.BOLD, size=14)
    progress = ft.ProgressBar(width=160, color=ft.Colors.BLUE_300, visible=False)
    log_lines: list[str] = []

    md_ctrls = [ddl_syntax, chk_assets, txt_asset]
    all_ctrls = [txt_input, txt_xml, txt_md, chk_empty, chk_skip, chk_assets,
                 ddl_syntax, txt_asset]

    btn_run = ft.Button(content=ft.Text("开始转换"), icon=ft.Icons.PLAY_ARROW,
                        bgcolor=ACC, color=ft.Colors.WHITE)
    btn_stop = ft.Button(content=ft.Text("停止"), icon=ft.Icons.STOP, visible=False,
                         bgcolor=ft.Colors.TRANSPARENT)

    # -- helpers ----------------------------------------------------
    def snack(msg: str):
        page.show_dialog(ft.SnackBar(ft.Text(msg), action="关闭",
                                     bgcolor=ft.Colors.ERROR_CONTAINER, open=True))

    def _push_log(msg: str):
        log_lines.append(msg)

    fp = ft.FilePicker()

    def browse_file(target: ft.TextField, _):
        async def _pick():
            r = await fp.pick_files(allowed_extensions=["one"])
            if r and r[0]:
                target.value = r[0].path
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
        for msg in log_lines[-200:]:  # show last 200 lines
            c = ft.Colors.ERROR if "[ERR]" in msg else None
            log_content.controls.append(
                ft.Text(msg, size=12, font_family="Consolas", selectable=True, color=c))

        def _close():
            try:
                page.pop_dialog()
            except Exception:
                pass

        footer = ("转换失败，请查看上方错误信息" if is_error
                  else "此窗口将在 2 秒后自动关闭")
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Icon(ft.Icons.TERMINAL_OUTLINED, color=ft.Colors.ERROR if is_error else ACC),
                ft.Text("运行日志" if not is_error else "运行日志 — 转换失败"),
            ], spacing=8),
            content=ft.Container(
                content=log_content,
                width=700, height=400,
                bgcolor=ft.Colors.SCRIM,
                border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
                border_radius=8,
                padding=12,
            ),
            actions=[ft.Text(footer, size=12, color=ft.Colors.ERROR if is_error else ft.Colors.SECONDARY)],
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

    # -- pipeline ---------------------------------------------------
    def run(_):
        nonlocal proc, running
        if running:
            return

        inp  = (txt_input.value or "").strip()
        xout = (txt_xml.value or "").strip() or str(ROOT / "output" / "xml")
        mout = (txt_md.value or "").strip() or str(ROOT / "output" / "markdown")
        adir = (txt_asset.value or "").strip()

        if not inp:              snack("请选择 .one 文件"); return
        if not os.path.isfile(inp): snack(f"未找到: {inp}"); return
        if not inp.lower().endswith(".one"): snack("必须是 .one 文件"); return
        if not chk_skip.value and chk_assets.value and not adir: snack("资源目录名不能为空"); return

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
        log_lines.append(f"输入: {inp}")
        log_lines.append(f"XML:  {xout}")
        if chk_skip.value:
            log_lines.append("Markdown: 跳过")
        else:
            log_lines.append(f"MD:   {mout}")
            log_lines.append(f"语法: {ddl_syntax.value}")
        log_lines.append("-" * 50)

        # Show popup log dialog
        show_log_dialog()

        args = [
            get_ps(), "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(PS1),
            "-InputOneFile", inp,
            "-XmlOutputDirectory", xout,
            "-MarkdownOutputDirectory", mout,
            "-LoadTimeoutSeconds", "30",
        ]
        if chk_empty.value:
            args.append("-IncludeEmptyPages")
        if chk_skip.value:
            args.append("-SkipMarkdownStage")
        if not chk_skip.value:
            args += ["-ImageSyntax", ddl_syntax.value or "markdown"]
            args += ["-AssetDirectoryName", adir]
            if not chk_assets.value:
                args.append("-CopyAssets:$false")

        def worker():
            nonlocal proc
            try:
                p = subprocess.Popen(
                    args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    text=True, encoding="utf-8", errors="replace",
                    cwd=str(ROOT),
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
                )
                proc = p
                out, err = p.communicate()
                if out:
                    for line in out.split("\n"):
                        if line.strip():
                            page.run_thread(_push_log, line.strip())
                if err:
                    for line in err.split("\n"):
                        if line.strip():
                            page.run_thread(_push_log, f"[ERR] {line.strip()}")
                page.run_thread(finish, p.returncode)
            except Exception as ex:
                page.run_thread(_push_log, f"[ERR] {ex}")
                page.run_thread(finish, -1)

        def finish(code: int):
            nonlocal running
            running = False
            btn_run.disabled = False
            btn_stop.visible = False
            if code == 0:
                status_text.value = "已完成"
                status_text.color = ft.Colors.GREEN_400
                progress.value = 1.0
                progress.color = ft.Colors.GREEN_300
            else:
                status_text.value = f"失败 (code={code})"
                status_text.color = ft.Colors.ERROR
                progress.value = 0
                progress.color = ft.Colors.ERROR
                log_lines.append("-" * 50)
                log_lines.append(f"[{ts()}] 结束, 退出码: {code}")
                show_log_dialog(is_error=True)
            for c in all_ctrls:
                c.disabled = False
            on_skip(None)
            page.update()

        threading.Thread(target=worker, daemon=True).start()

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

    # Auto-load default .one
    default = ROOT / "新临检.one"
    if default.exists():
        txt_input.value = str(default)


if __name__ == "__main__":
    ft.run(main, view=ft.AppView.FLET_APP)