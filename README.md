# 库存转导入模板脚本

这个项目把“药品库存.xlsx”转换成“饮片货位导入模板.xlsx”，规则与你当前使用的一致：

- 同一`饮片编码`允许有多个`批次`
- 导出时按`饮片编码`聚合库存（不同批次库存求和）
- 状态映射：`启用 -> 是`，`禁用 -> 否`
- `货位编号`默认 `Z999`
- `库存下限值`默认 `500`
- 输出前删除模板示例/旧数据行
- 默认按库存倒序排序（库存大的在上）

## 1) 安装 uv（一次）

Windows PowerShell:

```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

macOS / Linux:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

## 2) 初始化环境

在项目目录执行：

```bash
uv sync
```

## 3) 执行转换

```bash
uv run inventory-template-convert \
  --inventory "药品库存.xlsx" \
  --template "饮片货位导入模板.xlsx" \
  --output "饮片货位导入模板_最终.xlsx"
```

Windows 也可直接：

```powershell
uv run inventory-template-convert --inventory "药品库存.xlsx" --template "饮片货位导入模板.xlsx" --output "饮片货位导入模板_最终.xlsx"
```

## 4) 生成配方间回滚备份模板

这个命令会以`饮片货位导入模板.xlsx`为骨架（包括工作簿结构），
把`source`里的配方间库存源数据写成可导入的备份文件。

```bash
uv run source-backup-convert \
  --source "source/配方间饮片数据.xlsx" \
  --template "饮片货位导入模板.xlsx" \
  --output "source/饮片货位导入备份文件.xlsx"
```

Windows 也可一行执行：

```powershell
uv run source-backup-convert --source "source\\配方间饮片数据.xlsx" --template "饮片货位导入模板.xlsx" --output "source\\饮片货位导入备份文件.xlsx"
```

## 5) 一键执行（推荐）

你只要把文件放好：

- 根目录：库存文件 + 模板文件
- `source` 目录：`配方间饮片数据.xlsx`

然后执行：

Windows CMD:

```cmd
run_windows.cmd
```

macOS:

```bash
chmod +x run_mac.sh
./run_mac.sh
```

执行后会得到：

- 根目录：`饮片货位导入最终文件.xlsx`
- `source` 目录：`饮片货位导入备份文件.xlsx`

## 常用参数

- `--location` 货位编号，默认 `Z999`
- `--min-stock` 库存下限值，默认 `500`
- `--no-sort` 关闭库存倒序排序
- `--report-dir` 报告输出目录（默认当前目录）

示例：

```bash
uv run inventory-template-convert \
  --inventory "./data/new_inventory.xlsx" \
  --template "./template/饮片货位导入模板.xlsx" \
  --output "./out/import_template.xlsx" \
  --location "Z999" \
  --min-stock 500 \
  --report-dir "./out/reports"
```

## 输出内容

脚本会生成：

- 目标导入模板（你指定的 `--output`）
- `duplicate_code_batch_summary.csv`（重复的编码+批次汇总）
- `duplicate_code_batch_details.csv`（重复的编码+批次明细）
- `code_status_conflicts.csv`（同一编码状态冲突明细）
