#!/usr/bin/env python3
"""One-click: build exe → tag → push → GitHub Release with changelog."""

import os, re, subprocess, sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
EXE  = DIST / "OneConvert.exe"


def run(cmd: list[str], *, timeout: int | None = None):
    print(f"  > {' '.join(cmd)}")
    subprocess.run(cmd, cwd=str(ROOT), check=True,
                   encoding="utf-8", errors="replace",
                   timeout=timeout)


def out(cmd: list[str]) -> str:
    try:
        return subprocess.check_output(
            cmd, cwd=str(ROOT), text=True,
            encoding="utf-8", errors="replace",
        ).strip()
    except subprocess.CalledProcessError:
        return ""


def version_next() -> str:
    tags = [t.strip() for t in out(["git", "tag"]).split("\n") if t.strip()]
    for t in sorted(tags, reverse=True):
        m = re.match(r"^v(\d+)\.(\d+)\.(\d+)$", t)
        if m:
            return f"v{int(m[1])}.{int(m[2])}.{int(m[3]) + 1}"
    return "v1.0.0"


def changelog() -> str:
    prev = out(["git", "describe", "--tags", "--abbrev=0"])
    cmd = ["git", "log", "--pretty=format:- %s"]
    if prev:
        cmd.append(f"{prev}..HEAD")
    log = out(cmd)
    return log or "Initial release"


def build_exe():
    print("\n" + "=" * 60)
    print("  STEP 1/4  Build OneConvert.exe")
    print("=" * 60)
    sep = ";" if sys.platform == "win32" else ":"
    subprocess.run([
        sys.executable, "-m", "PyInstaller",
        "--onefile", "--windowed", "--name", "OneConvert",
        "--add-data", f"{ROOT / 'convert_onenote_xml.py'}{sep}.",
        "--add-data", f"{ROOT / 'Convert-OneNoteSectionToXml.ps1'}{sep}.",
        "--add-data", f"{ROOT / 'Convert-OneNoteToMarkdownPipeline.ps1'}{sep}.",
        "--collect-data", "flet",
        "--collect-data", "flet_core",
        "--collect-data", "flet_desktop",
        "--clean", "--noconfirm",
        str(ROOT / "OneConvertGUI.py"),
    ], cwd=str(ROOT), check=True, encoding="utf-8", errors="replace")
    if not EXE.exists():
        sys.exit(f"Build failed: {EXE} not found")
    print(f"  Done  ({EXE.stat().st_size / 1024 / 1024:.1f} MB)")


def stage_and_tag(version: str, changes: str):
    print("\n" + "=" * 60)
    print("  STEP 2/4  Git tag")
    print("=" * 60)
    run(["git", "add", "-A"])
    s = out(["git", "status", "--porcelain"])
    if s:
        run(["git", "commit", "-m", f"release: {version}"])
    else:
        print("  (nothing to commit)")
    short = changes.split("\n")[0][:80] if changes else version
    run(["git", "tag", "-a", version, "-m", short])

    print("\n" + "=" * 60)
    print("  STEP 3/4  Push")
    print("=" * 60)
    try:
        run(["git", "push", "origin", "HEAD"], timeout=30)
        run(["git", "push", "origin", version], timeout=30)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        print("  push 失败 (网络不通)，tag 已本地保存。")
        print(f"  稍后手动: git push origin HEAD && git push origin {version}")


def github_release(version: str, changes: str):
    print("\n" + "=" * 60)
    print("  STEP 4/4  GitHub Release")
    print("=" * 60)
    # Find gh CLI
    gh_paths = ["gh", r"C:\Program Files\GitHub CLI\gh.exe"]
    gh = None
    for p in gh_paths:
        try:
            if subprocess.run([p, "--version"], capture_output=True, timeout=5).returncode == 0:
                gh = p
                break
        except Exception:
            continue
    if not gh:
        print("  gh CLI 未安装，跳过 Release。")
        print(f"  手动: 在 GitHub Releases 创建 {version}，附加 {EXE.name}")
        return
        print("  gh CLI 未安装，跳过。")
        print(f"  手动: 在 GitHub Releases 创建 {version}，附加 {EXE.name}")
        return
    try:
        run([
            gh, "release", "create", version, str(EXE),
            "--title", f"OneConvert {version}",
            "--notes", changes,
        ], timeout=60)
        print(f"  Released {version}")
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        print(f"  上传失败 (网络不通)。")
        print(f"  手动: gh release create {version} {EXE}")


def main():
    os.chdir(str(ROOT))
    ver = version_next()
    log = changelog()

    print(f"\n  OneConvert Release — {ver}")
    print(f"  Changes:")
    for line in log.split("\n"):
        print(f"    {line}")

    build_exe()
    stage_and_tag(ver, log)
    github_release(ver, log)

    print(f"\n  {'=' * 60}")
    print(f"  完成 — {ver}")
    print(f"  Exe: {EXE}")
    print(f"  {'=' * 60}")


if __name__ == "__main__":
    main()
