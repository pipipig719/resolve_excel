from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

import openpyxl


INVENTORY_REQUIRED_HEADERS = ("饮片编码", "批次", "库存", "状态")
TEMPLATE_REQUIRED_HEADERS = ("饮片编码", "是否启用", "货位编号", "库存", "库存下限值")

STATUS_ENABLE_SRC = "启用"
STATUS_DISABLE_SRC = "禁用"
STATUS_ENABLE_OUT = "是"
STATUS_DISABLE_OUT = "否"


@dataclass
class InventoryRecord:
    row_num: int
    code: str
    batch: str
    stock: float
    status: str


@dataclass
class OutputRow:
    code: str
    enabled: str
    location: str
    stock: float | int
    min_stock: float | int


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def to_number(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).strip().replace(",", "")
    if raw == "":
        return 0.0
    try:
        return float(raw)
    except ValueError:
        return 0.0


def to_excel_number(value: float) -> float | int:
    rounded = round(value)
    if abs(value - rounded) < 1e-9:
        return int(rounded)
    return round(value, 6)


def read_header_map(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[str, int]:
    header_map: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        key = normalize_text(ws.cell(row=1, column=col).value)
        if key and key not in header_map:
            header_map[key] = col
    return header_map


def ensure_headers(header_map: dict[str, int], required: tuple[str, ...], file_label: str) -> None:
    missing = [h for h in required if h not in header_map]
    if missing:
        joined = ", ".join(missing)
        raise ValueError(f"{file_label} 缺少必需列: {joined}")


def load_inventory_records(inventory_path: Path) -> list[InventoryRecord]:
    wb = openpyxl.load_workbook(inventory_path, data_only=True)
    ws = wb.worksheets[0]
    header_map = read_header_map(ws)
    ensure_headers(header_map, INVENTORY_REQUIRED_HEADERS, f"库存文件 {inventory_path}")

    code_col = header_map["饮片编码"]
    batch_col = header_map["批次"]
    stock_col = header_map["库存"]
    status_col = header_map["状态"]

    records: list[InventoryRecord] = []
    for row in range(2, ws.max_row + 1):
        code = normalize_text(ws.cell(row=row, column=code_col).value)
        if code == "":
            continue
        batch = normalize_text(ws.cell(row=row, column=batch_col).value)
        stock = to_number(ws.cell(row=row, column=stock_col).value)
        status = normalize_text(ws.cell(row=row, column=status_col).value)
        records.append(
            InventoryRecord(
                row_num=row,
                code=code,
                batch=batch,
                stock=stock,
                status=status,
            )
        )
    return records


def map_status(status_set: set[str]) -> str:
    # 同一编码跨批次状态不一致时，优先按“启用”输出。
    if STATUS_ENABLE_SRC in status_set:
        return STATUS_ENABLE_OUT
    if STATUS_DISABLE_SRC in status_set:
        return STATUS_DISABLE_OUT
    return ""


def aggregate_records(
    records: list[InventoryRecord],
    location: str,
    min_stock: float | int,
) -> tuple[list[OutputRow], Counter[tuple[str, str]], dict[str, set[str]]]:
    pair_counter: Counter[tuple[str, str]] = Counter((r.code, r.batch) for r in records)
    code_stock: dict[str, float] = defaultdict(float)
    code_status: dict[str, set[str]] = defaultdict(set)
    code_order: list[str] = []
    seen_codes: set[str] = set()

    for r in records:
        code_stock[r.code] += r.stock
        if r.status:
            code_status[r.code].add(r.status)
        if r.code not in seen_codes:
            seen_codes.add(r.code)
            code_order.append(r.code)

    rows: list[OutputRow] = []
    for code in code_order:
        rows.append(
            OutputRow(
                code=code,
                enabled=map_status(code_status[code]),
                location=location,
                stock=to_excel_number(code_stock[code]),
                min_stock=min_stock,
            )
        )

    return rows, pair_counter, code_status


def write_template(
    template_path: Path,
    output_path: Path,
    rows: list[OutputRow],
    sort_desc: bool,
) -> None:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]
    header_map = read_header_map(ws)
    ensure_headers(header_map, TEMPLATE_REQUIRED_HEADERS, f"模板文件 {template_path}")

    if sort_desc:
        rows = sorted(rows, key=lambda x: float(x.stock), reverse=True)

    code_col = header_map["饮片编码"]
    enabled_col = header_map["是否启用"]
    location_col = header_map["货位编号"]
    stock_col = header_map["库存"]
    min_stock_col = header_map["库存下限值"]

    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    for row_idx, row_data in enumerate(rows, start=2):
        ws.cell(row=row_idx, column=code_col, value=row_data.code)
        ws.cell(row=row_idx, column=enabled_col, value=row_data.enabled)
        ws.cell(row=row_idx, column=location_col, value=row_data.location)
        ws.cell(row=row_idx, column=stock_col, value=row_data.stock)
        ws.cell(row=row_idx, column=min_stock_col, value=row_data.min_stock)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def write_reports(
    report_dir: Path,
    records: list[InventoryRecord],
    pair_counter: Counter[tuple[str, str]],
    code_status: dict[str, set[str]],
) -> tuple[Path, Path, Path]:
    report_dir.mkdir(parents=True, exist_ok=True)

    summary_file = report_dir / "duplicate_code_batch_summary.csv"
    details_file = report_dir / "duplicate_code_batch_details.csv"
    conflict_file = report_dir / "code_status_conflicts.csv"

    duplicate_pairs = sorted((k for k, count in pair_counter.items() if count > 1))

    with summary_file.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["code", "batch", "pair_count"])
        for code, batch in duplicate_pairs:
            writer.writerow([code, batch, pair_counter[(code, batch)]])

    with details_file.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["code", "batch", "pair_count", "source_row", "stock", "status"])
        for record in records:
            key = (record.code, record.batch)
            if key in duplicate_pairs:
                writer.writerow(
                    [
                        record.code,
                        record.batch,
                        pair_counter[key],
                        record.row_num,
                        record.stock,
                        record.status,
                    ]
                )

    with conflict_file.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["code", "statuses"])
        for code in sorted(code_status.keys()):
            statuses = {s for s in code_status[code] if s}
            if len(statuses) > 1:
                writer.writerow([code, "|".join(sorted(statuses))])

    return summary_file, details_file, conflict_file


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="将药品库存转换为饮片货位导入模板（按饮片编码聚合批次库存）"
    )
    parser.add_argument("--inventory", required=True, help="库存文件路径（xlsx）")
    parser.add_argument("--template", required=True, help="导入模板路径（xlsx）")
    parser.add_argument(
        "--output",
        help="输出文件路径（xlsx）。不传时默认在模板同目录生成 *_generated.xlsx",
    )
    parser.add_argument(
        "--location",
        default="Z999",
        help="货位编号默认值，默认 Z999",
    )
    parser.add_argument(
        "--min-stock",
        type=float,
        default=500,
        help="库存下限值默认值，默认 500",
    )
    parser.add_argument(
        "--no-sort",
        action="store_true",
        help="不按库存倒序排序（默认按库存从大到小排序）",
    )
    parser.add_argument(
        "--report-dir",
        default=".",
        help="报告文件输出目录，默认当前目录",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    inventory_path = Path(args.inventory).expanduser().resolve()
    template_path = Path(args.template).expanduser().resolve()

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = template_path.with_name(f"{template_path.stem}_generated{template_path.suffix}")

    report_dir = Path(args.report_dir).expanduser().resolve()

    if not inventory_path.exists():
        print(f"[ERROR] 库存文件不存在: {inventory_path}", file=sys.stderr)
        return 1
    if not template_path.exists():
        print(f"[ERROR] 模板文件不存在: {template_path}", file=sys.stderr)
        return 1

    min_stock_value: float | int = args.min_stock
    if abs(args.min_stock - round(args.min_stock)) < 1e-9:
        min_stock_value = int(round(args.min_stock))

    try:
        records = load_inventory_records(inventory_path)
        rows, pair_counter, code_status = aggregate_records(
            records=records,
            location=args.location,
            min_stock=min_stock_value,
        )
        write_template(
            template_path=template_path,
            output_path=output_path,
            rows=rows,
            sort_desc=not args.no_sort,
        )
        summary_file, details_file, conflict_file = write_reports(
            report_dir=report_dir,
            records=records,
            pair_counter=pair_counter,
            code_status=code_status,
        )
    except Exception as exc:  # noqa: BLE001
        print(f"[ERROR] 处理失败: {exc}", file=sys.stderr)
        return 1

    duplicate_pair_count = sum(1 for _, count in pair_counter.items() if count > 1)
    duplicate_pair_rows = sum(count for _, count in pair_counter.items() if count > 1)

    print("[OK] 转换完成")
    print(f"inventory: {inventory_path}")
    print(f"template:  {template_path}")
    print(f"output:    {output_path}")
    print(f"source_rows: {len(records)}")
    print(f"unique_codes: {len(rows)}")
    print(f"duplicate_code_batch_count: {duplicate_pair_count}")
    print(f"duplicate_code_batch_rows:  {duplicate_pair_rows}")
    print(f"report_summary: {summary_file}")
    print(f"report_details: {details_file}")
    print(f"report_status_conflicts: {conflict_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
