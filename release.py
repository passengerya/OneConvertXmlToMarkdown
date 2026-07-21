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
    """Read the highest tag that exists on the REMOTE (only pushed tags count).
    Falls back to local tags if network unreachable."""
    try:
        run(["git", "fetch", "--tags", "origin"], timeout=15, check=False)
    except Exception:
        pass
    # Prefer remote tags
    raw = out(["git", "ls-remote", "--tags", "origin"])
    tags: list[str] = []
    for line in raw.split("\n"):
        if "refs/tags/" in line and "^{}" not in line:
            t = line.strip().split("refs/tags/")[-1]
            if t:
                tags.append(t)
    # Fallback to local tags if remote unreachable
    if not tags:
        local = [t.strip() for t in out(["git", "tag"]).split("\n") if t.strip()]
        tags = local
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
    """Commit → push → tag (on success) → GitHub Release."""
    if not EXE.exists():
        sys.exit(f"  exe 不存在，请先选择 [1] 打包。\n  期望路径: {EXE}")

    ver = version_next()
    log = changelog()

    print(f"\n  版本: {ver}")
    print(f"  变更:")
    for line in log.split("\n")[:15]:
        print(f"    {line}")

    # Commit (no tag yet — tag only after successful push)
    print("\n" + "-" * 40)
    print("  Git commit")
    print("-" * 40)
    run(["git", "add", "-A"])
    s = out(["git", "status", "--porcelain"])
    if s:
        run(["git", "commit", "-m", f"release: {ver}"])
    else:
        print("  (nothing to commit)")

    # Push HEAD
    print("\n" + "-" * 40)
    print("  Push HEAD")
    print("-" * 40)
    try:
        run(["git", "push", "origin", "HEAD"], timeout=30)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        print("  !! push 失败 (网络不通)，版本号不递增。")
        print(f"  修复网络后重试，仍为 {ver}")
        return

    # Push succeeded → create tag and push it
    print("\n" + "-" * 40)
    print("  Push tag")
    print("-" * 40)
    # Clean stale local tag from any previous failed push
    subprocess.run(["git", "tag", "-d", ver], capture_output=True, cwd=str(ROOT))
    short = log.split("\n")[0][:80] if log else ver
    run(["git", "tag", "-a", ver, "-m", short])
    try:
        run(["git", "push", "origin", ver], timeout=30)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        print(f"  !! tag push 失败，但 HEAD 已推送。手动: git push origin {ver}")
        return

    # GitHub Release
    print("\n" + "-" * 40)
    print("  GitHub Release")
    print("-" * 40)
    gh = find_gh()
    if not gh:
        print("  gh CLI 未安装，跳过。")
        print(f"  手动: gh release create {ver} {EXE}")
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
