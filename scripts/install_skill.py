"""Install this skill without overwriting an existing installation."""
from __future__ import annotations

import argparse
import os
import shutil
from pathlib import Path


def main():
    codex_dir = Path(os.environ.get("CODEX_HOME", str(Path.home() / ".codex")))
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--skills-dir", type=Path, default=codex_dir / "skills")
    args = parser.parse_args()
    source = Path(__file__).resolve().parents[1]
    destination = args.skills_dir.expanduser().resolve() / "editable-visio-recreation"
    if not (source / "SKILL.md").is_file():
        parser.error("Run this installer from a complete repository or skill folder.")
    if destination == source or source in destination.parents:
        parser.error("The destination must be outside the source repository.")
    if destination.exists() or destination.is_symlink():
        parser.error("An installation already exists; it was left unchanged: " + str(destination))
    destination.mkdir(parents=True)
    try:
        shutil.copy2(source / "SKILL.md", destination / "SKILL.md")
        shutil.copy2(source / "LICENSE", destination / "LICENSE")
        for folder in ("agents", "scripts", "references"):
            shutil.copytree(
                source / folder, destination / folder,
                ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
            )
    except OSError as exc:
        parser.exit(1, "Installation incomplete at " + str(destination) + ": " + str(exc) + "\n")
    print("Installed editable-visio-recreation at " + str(destination))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

