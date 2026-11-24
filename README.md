# Excel Tools（Python Excel 效率工具集）

> 原始开发时间：2025-07 ~ 2025-08

集合两个用于日常办公表格处理的 Python 桌面 GUI 工具，均支持通过 PyInstaller 打包为 Windows 可执行程序。

## 包含工具

| 子目录 | 名称 | 原始时间 | 主要作用 |
| --- | --- | --- | --- |
| [duplicate-checker/](./duplicate-checker/README.md) | Excel 重复数据检测工具 | 2025-07-24 | 检测表格指定列中因多余空格导致的隐藏重复项，支持报告导出与一键清理 |
| [table-filler/](./table-filler/README.md) | Excel 表格填写工具 | 2025-08 | 按模板格式自动匹配并填写发货明细表，支持一对多匹配与列映射配置保存（原 Gitee 仓库：`gengu0312/excel-fill-tools`） |

## 使用说明

进入对应子目录安装依赖并运行：

```bash
pip install -r requirements.txt
```
