from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

FINAL_OUTPUT_NAME = "\u6000\u5b81\u996e\u7247\u8d27\u4f4d\u5bfc\u5165\u6700\u7ec8\u6587\u4ef6.xlsx"
BACKUP_OUTPUT_NAME = "\u6000\u5b81\u996e\u7247\u8d27\u4f4d\u5bfc\u5165\u5907\u4efd\u6587\u4ef6.xlsx"


def find_first_file(directory: Path, exts: tuple[str, ...], exclude_keywords: tuple[str, ...] = ()) -> Path:
    candidates: list[Path] = []
    for p in directory.iterdir():
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in exts:
            continue
        if any(k in p.name for k in exclude_keywords):
            continue
        candidates.append(p)

    if not candidates:
        raise FileNotFoundError(f"No file found in {directory} with extensions {exts}")

    candidates.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return candidates[0]


def find_template_file(project_root: Path, huaining_dir: Path) -> Path:
    root_files = [p for p in project_root.glob("*.xlsx") if p.is_file() and not p.name.startswith("~$")]
    root_preferred = [p for p in root_files if "\u6a21\u677f" in p.name]
    if root_preferred:
        root_preferred.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return root_preferred[0]
    if root_files:
        root_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return root_files[0]

    local_files = [p for p in huaining_dir.glob("*.xlsx") if p.is_file() and not p.name.startswith("~$")]
    local_preferred = [p for p in local_files if "\u6a21\u677f" in p.name]
    if local_preferred:
        local_preferred.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return local_preferred[0]
    if local_files:
        local_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return local_files[0]

    raise FileNotFoundError("No template xlsx found in project root or huaining directory")


def run_cmd(command: list[str], cwd: Path) -> None:
    result = subprocess.run(command, cwd=cwd, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"Command failed (exit={result.returncode}): {' '.join(command)}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Huaining inventory converter")
    parser.add_argument("--inventory", help="Huaining inventory file path")
    parser.add_argument("--source", help="Huaining source file path")
    parser.add_argument("--template", help="Template xlsx file path")
    parser.add_argument("--output-final", help="Final output xlsx path")
    parser.add_argument("--output-backup", help="Backup output xlsx path")
    parser.add_argument("--report-dir", help="Report output dir path")
    parser.add_argument("--keep-reports", action="store_true", help="Keep report files")
    parser.add_argument("--default-location", default="Z999")
    parser.add_argument("--default-min-stock", type=float, default=500)
    parser.add_argument("--no-sort", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    huaining_dir = Path(__file__).resolve().parent
    project_root = huaining_dir.parent
    source_dir = huaining_dir / "huaining_source"

    try:
        if args.inventory:
            inventory_file = Path(args.inventory).resolve()
        else:
            inventory_file = find_first_file(
                huaining_dir,
                (".xlsx", ".xls"),
                (
                    "\u5907\u4efd\u6587\u4ef6",
                    "\u6700\u7ec8\u6587\u4ef6",
                    "\u6a21\u677f",
                    "process_huaining",
                    "run_huaining",
                ),
            )

        if args.source:
            source_file = Path(args.source).resolve()
        else:
            source_file = find_first_file(source_dir, (".xls", ".xlsx"), ("\u5907\u4efd\u6587\u4ef6",))

        template_file = Path(args.template).resolve() if args.template else find_template_file(project_root, huaining_dir)

        final_out = (
            Path(args.output_final).resolve()
            if args.output_final
            else (huaining_dir / FINAL_OUTPUT_NAME).resolve()
        )
        backup_out = (
            Path(args.output_backup).resolve()
            if args.output_backup
            else (source_dir / BACKUP_OUTPUT_NAME).resolve()
        )
        report_dir = (
            Path(args.report_dir).resolve()
            if args.report_dir
            else (huaining_dir / ".tmp_reports").resolve()
        )

        convert_inventory_cmd = [
            sys.executable,
            "convert_inventory.py",
            "--inventory",
            str(inventory_file),
            "--template",
            str(template_file),
            "--source-for-match",
            str(source_file),
            "--output",
            str(final_out),
            "--location",
            args.default_location,
            "--min-stock",
            str(args.default_min_stock),
            "--report-dir",
            str(report_dir),
        ]
        if args.no_sort:
            convert_inventory_cmd.append("--no-sort")

        convert_backup_cmd = [
            sys.executable,
            "convert_source_backup.py",
            "--source",
            str(source_file),
            "--template",
            str(template_file),
            "--output",
            str(backup_out),
            "--default-location",
            args.default_location,
            "--default-min-stock",
            str(args.default_min_stock),
        ]

        run_cmd(convert_inventory_cmd, cwd=project_root)
        run_cmd(convert_backup_cmd, cwd=project_root)

        if not args.keep_reports and report_dir.exists():
            shutil.rmtree(report_dir, ignore_errors=True)

        print("[OK] Huaining conversion complete")
        print(f"inventory_file={inventory_file}")
        print(f"source_file={source_file}")
        print(f"template_file={template_file}")
        print(f"final_output={final_out}")
        print(f"backup_output={backup_out}")
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
