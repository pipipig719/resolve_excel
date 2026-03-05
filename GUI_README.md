# GUI Launcher (Windows + macOS)

This GUI lets you select or drag files and then generate:

- `饮片货位导入最终文件.xlsx` (in selected output directory)
- `source/饮片货位导入备份文件.xlsx` (under selected output directory)

## Files

- `gui_launcher.py`
- `run_gui_windows.cmd`
- `run_gui_mac.sh`

## Run

Windows (CMD):

```cmd
run_gui_windows.cmd
```

macOS:

```bash
chmod +x run_gui_mac.sh
./run_gui_mac.sh
```

## Inputs in GUI

- Inventory file: `.xlsx`
- Pharmacy-room source file: `.xls` or `.xlsx`
- Template file: `.xlsx`
- Output directory: where final/backup files are written

## Drag & Drop

Drag & drop is enabled when `tkinterdnd2` is available.
If drag does not work, click `Browse` buttons (always supported).
