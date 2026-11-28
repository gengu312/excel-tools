# Excel Tools (Python Excel Office Utilities)

<p>
  <a href="./README.md"><img src="https://img.shields.io/badge/Docs-English-2563EB?style=flat-square" alt="English" /></a>
  <a href="./README.zh-CN.md"><img src="https://img.shields.io/badge/%E6%96%87%E6%A1%A3-%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87-64748B?style=flat-square" alt="简体中文" /></a>
</p>

> Original Development Date: 2025-07 ~ 2025-08

A collection of two Python desktop GUI utilities for everyday spreadsheet processing, both supporting packaging into standalone Windows executables via PyInstaller.

## Included Tools

| Subdirectory | Name | Original Date | Description |
| --- | --- | --- | --- |
| [duplicate-checker/](./duplicate-checker/README.md) | Excel Duplicate Data Checker | 2025-07-24 | Detects hidden duplicate entries caused by trailing/extra spaces in specified columns, with report export and one-click cleanup |
| [table-filler/](./table-filler/README.md) | Excel Template Auto-Filler | 2025-08 | Automatically matches and populates shipment detail records into template spreadsheets (one-to-many), with persistent column-mapping profiles |

## Usage

Navigate into the desired tool subdirectory and install dependencies before running:

```bash
pip install -r requirements.txt
```
