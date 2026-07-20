#!/usr/bin/env python3
"""OneConvert release tool — build exe / push release / both."""

import os, re, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
EXE  = DIST / "OneConvert.exe"


def run(cmd: list[str], *, timeout: int | None = None, check: bool = True):
    print(f"  > {' '.join(cmd)}")
    return subprocess.run(cmd, cwd=str(ROOT),
                          encoding="utf-8", errors="replace",
                          timeout=timeout, check=check)


def out(cmd: list[str]) -> str:
    try:
        return subprocess.check_output(
            cmd, cwd=str(ROOT), text=True,
            encoding="utf-8", errors="replace",
        ).strip()
    except subprocess.CalledProcessError:
        return ""


def version_next() -> str:
    # fetch remote tags so we see what's actually released
    try:
        run(["git", "fetch", "--tags", "origin"], timeout=15, check=False)
    except Exception:
        pass
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


def find_gh() -> str | None:
    for p in ["gh", r"C:\Program Files\GitHub CLI\gh.exe"]:
        try:
            if subprocess.run([p, "--version"], capture_output=True, timeout=5).returncode == 0:
                return p
        except Exception:
            continue
    return None


# ═══════════════════════════════════════════════════════════════
#  Actions
# ═══════════════════════════════════════════════════════════════

def action_build():
    print("\n" + "=" * 60)
    print("  打包 OneConvert.exe (PyInstaller)")
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
        sys.exit("Build failed: exe not found")
    print(f"\n  打包完成  ({EXE.stat().st_size / 1024 / 1024:.1f} MB)")


def action_release():
    """Tag, commit, push, and create GitHub Release.  Fails if no exe exists."""
    if not EXE.exists():
        sys.exit(f"  exe 不存在，请先选择 [1] 打包。\n  期望路径: {EXE}")

    ver = version_next()
    log = changelog()

    print(f"\n  版本: {ver}")
    print(f"  变更:")
    for line in log.split("\n")[:15]:
        print(f"    {line}")

    # Commit & tag
    print("\n" + "-" * 40)
    print("  Git commit + tag")
    print("-" * 40)
    run(["git", "add", "-A"])
    s = out(["git", "status", "--porcelain"])
    if s:
        run(["git", "commit", "-m", f"release: {ver}"])
    else:
        print("  (nothing to commit)")
    short = log.split("\n")[0][:80] if log else ver
    run(["git", "tag", "-a", ver, "-m", short])

    # Push
    print("\n" + "-" * 40)
    print("  Push")
    print("-" * 40)
    pushed = True
    try:
        run(["git", "push", "origin", "HEAD"], timeout=30)
        run(["git", "push", "origin", ver], timeout=30)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        print("  !! push 失败 (网络不通)")
        print(f"  手动: git push origin HEAD && git push origin {ver}")
        pushed = False

    # GitHub Release
    print("\n" + "-" * 40)
    print("  GitHub Release")
    print("-" * 40)
    gh = find_gh()
    if not gh:
        print("  gh CLI 未安装，跳过。")
        print(f"  手动: gh release create {ver} {EXE}")
        return

    if not pushed:
        print(f"  push 未完成，跳过 Release。先 push 后再创建。")
        return

    try:
        run([gh, "release", "create", ver, str(EXE),
             "--title", f"OneConvert {ver}",
             "--notes", log], timeout=60)
        print(f"\n  Released {ver} ✓")
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        print(f"  !! Release 创建失败 (网络不通)")
        print(f"  手动: gh release create {ver} {EXE}")


def action_build_and_release():
    action_build()
    action_release()


# ═══════════════════════════════════════════════════════════════
#  Menu
# ═══════════════════════════════════════════════════════════════

MENU = {
    "1": ("打包 (仅构建 exe)", action_build),
    "2": ("推送 Release (tag + push + GitHub Release)", action_release),
    "3": ("打包并推送 Release", action_build_and_release),
}


def main():
    os.chdir(str(ROOT))
    ver = version_next()

    print(f"\n  OneConvert Release Tool")
    print(f"  当前最新版本: {ver}")
    print(f"  选择操作:")
    for k, (desc, _) in MENU.items():
        print(f"    [{k}] {desc}")

    choice = input("\n  输入序号 (1/2/3): ").strip()
    action = MENU.get(choice)
    if not action:
        print("  无效选项")
        sys.exit(1)

    action[1]()
    print()


if __name__ == "__main__":
    main()
