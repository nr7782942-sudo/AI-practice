#!/usr/bin/env python3
"""Offline, no-overwrite installer. Requires Python 3.10 or newer."""
from __future__ import annotations

import argparse
import ctypes
import hashlib
import json
import os
from pathlib import Path
import stat
import sys
from datetime import datetime, timezone

PAYLOAD = ("CLAUDE.md", ".claude", "scripts", "docs")
GENERATED = ("dram-config.json", "INSTALL-JOURNAL.jsonl", "INSTALL-RECEIPT.json")


def stamp() -> str:
    return datetime.now(timezone.utc).isoformat()


def reject_link(path: Path) -> None:
    """Check every existing ancestor, including Windows junction/reparse paths."""
    for item in reversed((path, *path.parents)):
        try:
            info = item.lstat()
        except FileNotFoundError:
            continue
        if stat.S_ISLNK(info.st_mode) or getattr(info, "st_file_attributes", 0) & 0x400:
            raise ValueError(f"Symlink/junction/reparse path is not allowed: {item}")


def local_path(value: str) -> Path:
    if value.startswith(("\\\\", "//")) or ".." in Path(value).parts:
        raise ValueError("UNC/device paths and '..' path components are not allowed.")
    path = Path(os.path.abspath(value))
    reject_link(path)
    if os.name == "nt":
        if not path.drive or ":" in str(path)[2:]:
            raise ValueError(f"Invalid local Windows path: {path}")
        for part in path.parts[1:]:
            if part.rstrip(" .") != part or Path(part).is_reserved():
                raise ValueError(f"Ambiguous/reserved Windows path component: {part}")
        drive_type = ctypes.windll.kernel32.GetDriveTypeW(str(path.anchor))
        if drive_type == 4:
            raise ValueError(f"Network-mapped drives are not allowed: {path.drive}")
    return path


def overlaps(a: Path, b: Path) -> bool:
    return a == b or a in b.parents or b in a.parents


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def inventory(package: Path) -> tuple[list[Path], list[Path]]:
    files: list[Path] = []
    directories: list[Path] = []

    def visit(path: Path) -> None:
        reject_link(path)
        rel = path.relative_to(package)
        if any(part == ".obsidian" for part in rel.parts) or path.name.lower() == "settings.json":
            raise ValueError(f"Forbidden configuration in package: {rel}")
        if path.is_dir():
            directories.append(rel)
            for child in sorted(path.iterdir()):
                visit(child)
        elif path.is_file():
            files.append(rel)
        else:
            raise ValueError(f"Unsupported payload entry: {rel}")

    for name in PAYLOAD:
        src = package / name
        reject_link(src)
        if not src.exists():
            raise ValueError(f"Incomplete package: missing {name}")
        visit(src)
    return sorted(set(directories), key=lambda p: (len(p.parts), str(p))), sorted(files)


def exclusive_write(path: Path, content: bytes) -> None:
    reject_link(path.parent)
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_BINARY", 0)
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    descriptor = os.open(str(path), flags, 0o600)
    with os.fdopen(descriptor, "wb") as handle:
        handle.write(content)
        handle.flush()
        os.fsync(handle.fileno())


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def install(package: Path, project: Path, vault: Path, dry_run: bool) -> None:
    if overlaps(package, project) or overlaps(package, vault):
        raise ValueError("The extracted package must be separate from the project and vault.")
    if not vault.is_dir():
        raise ValueError(f"Vault must be an existing local directory: {vault}")
    if project.exists() and not project.is_dir():
        raise ValueError(f"Project is not a directory: {project}")
    directories, files = inventory(package)
    collisions = []
    for rel in [*directories, *files, *(Path(name) for name in GENERATED)]:
        dest = project / rel
        reject_link(dest)
        if dest.exists() and (rel not in directories or not dest.is_dir()):
            collisions.append(str(rel))
    if collisions:
        raise ValueError("No files were changed. Existing target files/conflicts:\n  " + "\n  ".join(collisions))
    # Hash/read every payload file before any mutation; detects unreadable payloads early.
    expected = {str(rel): digest(package / rel) for rel in files}
    print(f"Project: {project}\nVault (read-only during installation): {vault}")
    print(f"Payload files: {len(files)}; generated files: {', '.join(GENERATED)}")
    if dry_run:
        for rel in files:
            print(f"  CREATE {rel}")
        print("DRY RUN: no directories or files created.")
        return
    reject_link(project)
    project.mkdir(parents=True, exist_ok=True)
    journal = project / "INSTALL-JOURNAL.jsonl"
    exclusive_write(journal, (json.dumps({"event": "started", "time": stamp(), "project": str(project), "vault": str(vault)}, ensure_ascii=False) + "\n").encode("utf-8"))
    installed = []

    def record(event: str, **details: object) -> None:
        reject_link(journal)
        with journal.open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps({"event": event, "time": stamp(), **details}, ensure_ascii=False) + "\n")
            handle.flush()
            os.fsync(handle.fileno())

    try:
        for rel in directories:
            reject_link(project / rel)
            (project / rel).mkdir(exist_ok=True)
        for rel in files:
            source = package / rel
            reject_link(source)
            content = source.read_bytes()
            actual = hashlib.sha256(content).hexdigest()
            if actual != expected[str(rel)]:
                raise ValueError(f"Payload changed after preflight: {rel}")
            exclusive_write(project / rel, content)
            entry = {"path": rel.as_posix(), "sha256": actual, "bytes": len(content)}
            installed.append(entry)
            record("created", **entry)
        config = json_bytes({"schema_version": 1, "vault_path": str(vault)})
        exclusive_write(project / "dram-config.json", config)
        entry = {"path": "dram-config.json", "sha256": hashlib.sha256(config).hexdigest(), "bytes": len(config)}
        installed.append(entry)
        record("created", **entry)
        record("payload_complete", receipt="INSTALL-RECEIPT.json")
        installed.append({"path": journal.name, "sha256": digest(journal), "bytes": journal.stat().st_size})
        receipt = {"schema_version": 1, "status": "complete", "installed_at": stamp(), "project_path": str(project), "vault_path": str(vault), "files": installed, "note": "Receipt excludes its own hash. Vault was not modified."}
        exclusive_write(project / "INSTALL-RECEIPT.json", json_bytes(receipt))
    except Exception as exc:
        try:
            record("failed", error=str(exc), completed_files=len(installed))
        except Exception:
            pass
        raise RuntimeError(f"Installation stopped; created files were retained. Review {journal}. Retry into a new empty project. Cause: {exc}") from exc
    print(f"Installation complete. Receipt: {project / 'INSTALL-RECEIPT.json'}")


def main() -> int:
    # Redirected Windows CMD output may use a legacy code page; path display must not abort installation.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="backslashreplace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project", help="Destination project directory (existing or new)")
    parser.add_argument("--vault", help="Existing local Obsidian vault directory")
    parser.add_argument("--dry-run", action="store_true", help="Validate and list changes without writing")
    args = parser.parse_args()
    if sys.version_info < (3, 10):
        parser.error("Python 3.10 or newer is required.")
    if not args.project or not args.vault:
        parser.print_help()
        print("\nNothing changed. Supply both --project and --vault.", file=sys.stderr)
        return 2
    try:
        package = local_path(str(Path(__file__).absolute().parent))
        install(package, local_path(args.project), local_path(args.vault), args.dry_run)
    except (OSError, ValueError, RuntimeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
