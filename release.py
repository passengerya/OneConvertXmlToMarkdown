#!/usr/bin/env python3
"""One-click: build exe → tag → GitHub Release with changelog."""

import os, re, subprocess, sys, textwrap
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DIST = ROOT / "dist"
EXE  = DIST / "OneConvert.exe"


def run(cmd: list[str], **kw) -> subprocess.CompletedProcess:
    print(f"  $ {' '.join(cmd)}")
    kw.setdefault("encoding", "utf-8")
    kw.setdefault("errors", "replace")
    return subprocess.run(cmd, cwd=str(ROOT), check=True, **kw)


def get_version() -> str:
    """Return next version from existing tags, or use date-based default."""
    try:
        tags = subprocess.check_output(
            ["git", "tag", "--sort=-v:refname"],
            cwd=str(ROOT), text=True, encoding="utf-8", errors="replace",
        ).strip().split("\n")
    except Exception:
        tags = []
    tags = [t.strip() for t in tags if t.strip()]
    for tag in tags:
        m = re.match(r"^v(\d+)\.(\d+)\.(\d+)$", tag)
        if m:
            major, minor, patch = map(int, m.groups())
            return f"v{major}.{minor}.{patch + 1}"
    return f"v{datetime.now().strftime('%y')}.1.0"  # e.g. v25.1.0


def get_changelog(last_tag: str | None) -> str:
    """Extract commit messages since last tag."""
    cmd = ["git", "log", "--pretty=format:- %s"]
    if last_tag:
        cmd.append(f"{last_tag}..HEAD")
    try:
        log = subprocess.check_output(
            cmd, cwd=str(ROOT), text=True, encoding="utf-8", errors="replace",
        ).strip()
    except subprocess.CalledProcessError:
        log = ""
    return log or "Initial release"


def build_exe() -> bool:
    """Run PyInstaller. Return True on success."""
    print("\n" + "=" * 60)
    print("  STEP 1/4 — Building OneConvert.exe")
    print("=" * 60)

    sep = ";" if sys.platform == "win32" else ":"

    args = [
        sys.executable, "-m", "PyInstaller",
        "--onefile", "--windowed",
        "--name", "OneConvert",
        "--add-data", f"{ROOT / 'convert_onenote_xml.py'}{sep}.",
        "--add-data", f"{ROOT / 'Convert-OneNoteSectionToXml.ps1'}{sep}.",
        "--add-data", f"{ROOT / 'Convert-OneNoteToMarkdownPipeline.ps1'}{sep}.",
        "--hidden-import", "flet",
        "--hidden-import", "flet_core",
        "--hidden-import", "flet_desktop",
        "--clean",
        "--noconfirm",
        str(ROOT / "OneConvertGUI.py"),
    ]

    result = subprocess.run(args, cwd=str(ROOT),
                            encoding="utf-8", errors="replace")
    if result.returncode != 0:
        print("\n  !! Build failed")
        return False
    if not EXE.exists():
        print(f"\n  !! EXE not found at {EXE}")
        return False
    print(f"\n  Built: {EXE}  ({EXE.stat().st_size / 1024 / 1024:.1f} MB)")
    return True


def main():
    os.chdir(str(ROOT))

    # ── check prerequisites ──────────────────────────────
    missing = []
    for cmd, name in [("powershell", "PowerShell"), ("python", "Python"),
                       ("git", "git"), ("gh", "gh CLI")]:
        found = subprocess.run(["where", cmd], capture_output=True, shell=True,
                               encoding="utf-8", errors="replace")
        if found.returncode != 0:
            missing.append(name)
    if missing:
        print(f"  WARNING: not found: {', '.join(missing)}")

    # ── determine version ─────────────────────────────────
    version = get_version()
    try:
        last_tag = subprocess.check_output(
            ["git", "describe", "--tags", "--abbrev=0"],
            cwd=str(ROOT), text=True, encoding="utf-8", errors="replace",
        ).strip()
    except Exception:
        last_tag = None

    changelog = get_changelog(last_tag)

    print("\n" + "=" * 60)
    print(f"  OneConvert Release — {version}")
    print("=" * 60)
    print(f"\n  Previous tag: {last_tag or '(none)'}")
    print(f"\n  Changes:")
    for line in changelog.split("\n")[:20]:
        print(f"    {line}")
    if len(changelog.split("\n")) > 20:
        print(f"    ... and {len(changelog.split(changelog)) - 20} more")
    print()

    # ── step 1: build ─────────────────────────────────────
    if not build_exe():
        sys.exit(1)

    # ── step 2: commit pending changes ────────────────────
    print("\n" + "=" * 60)
    print("  STEP 2/4 — Staging changes")
    print("=" * 60)
    try:
        run(["git", "add", "-A"])
        status = subprocess.check_output(
            ["git", "status", "--porcelain"], cwd=str(ROOT), text=True,
            encoding="utf-8", errors="replace",
        ).strip()
        if status:
            run(["git", "commit", "-m", f"release: {version}"])
        else:
            print("  (nothing to commit)")
    except subprocess.CalledProcessError as e:
        print(f"  Warning: {e}")

    # ── step 3: tag ───────────────────────────────────────
    print("\n" + "=" * 60)
    print("  STEP 3/4 — Creating tag & pushing")
    print("=" * 60)
    run(["git", "tag", "-a", version, "-m", f"{version}: {changelog.split(chr(10))[0]}"])
    run(["git", "push", "origin", version])

    # ── step 4: GitHub Release ────────────────────────────
    print("\n" + "=" * 60)
    print("  STEP 4/4 — Creating GitHub Release")
    print("=" * 60)
    run([
        "gh", "release", "create", version,
        str(EXE),
        "--title", f"OneConvert {version}",
        "--notes", changelog,
    ])

    print("\n" + "=" * 60)
    print(f"  Released {version} successfully!")
    print(f"  {EXE}")
    print("=" * 60)


if __name__ == "__main__":
    main()
