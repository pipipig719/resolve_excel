# GUI Launcher (Windows + macOS)

This GUI supports two modes:

- `怀宁`
- `肥西`

Pick a mode, then select or drag files and generate:

- Final import workbook (mode-specific name)
- Backup import workbook (mode-specific name)

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

- Processing mode: `怀宁` or `肥西`
- Inventory file: `.xlsx` / `.xls` (mode-dependent)
- Pharmacy-room source file: `.xls` or `.xlsx`
- Template file: `.xlsx`
- Output directory: final + backup files are generated here

Backup file is generated under:

- `huaining_source/` in `怀宁` mode
- `feixi_source/` in `肥西` mode

## Drag & Drop

Drag & drop is enabled when `tkinterdnd2` is available.
If drag does not work, use the `Browse` buttons.
