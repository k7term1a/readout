#!/usr/bin/env python3
"""Zip every folder under skills/ into dist/<name>.skill (the format claude.ai accepts for upload)."""
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKIP = {"__pycache__", ".DS_Store"}


def package(skill_dir: Path, out_dir: Path) -> Path:
    if not (skill_dir / "SKILL.md").exists():
        raise SystemExit(f"{skill_dir} has no SKILL.md")
    out = out_dir / f"{skill_dir.name}.skill"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(skill_dir.rglob("*")):
            if f.is_file() and not SKIP.intersection(f.parts) and f.suffix != ".pyc":
                z.write(f, f.relative_to(skill_dir.parent))
    return out


def main():
    out_dir = ROOT / "dist"
    out_dir.mkdir(exist_ok=True)
    names = sys.argv[1:] or [p.name for p in (ROOT / "skills").iterdir() if p.is_dir()]
    for n in names:
        print("packaged", package(ROOT / "skills" / n, out_dir).relative_to(ROOT))


if __name__ == "__main__":
    main()
