# Excel Tools（Python Excel 办公工具集）

<p>
  <a href="./README.md"><img src="https://img.shields.io/badge/Docs-English-64748B?style=flat-square" alt="English" /></a>
  <a href="./README.zh-CN.md"><img src="https://img.shields.io/badge/%E6%96%87%E6%A1%A3-%E7%AE%80%E4%BD%93%E4%B8%AD%E6%96%87-2563EB?style=flat-square" alt="简体中文" /></a>
</p>

> 原始开发时间：2025-07 ~ 2025-08

集合两个用于日常办公表格处理的 Python 桌面 GUI 工具，均支持通过 PyInstaller 打包为 Windows 可执行程序。

## 包含工具

| 子目录 | 名称 | 原始时间 | 主要作用 |
| --- | --- | --- | --- |
| [duplicate-checker/](./duplicate-checker/README.md) | Excel 重复数据检测工具 | 2025-07-24 | 检测表格指定列中因多余空格导致的隐藏重复项，支持报告导出与一键清理 |
| [table-filler/](./table-filler/README.md) | Excel 表格填写工具 | 2025-08 | 将发货明细表数据按模板表格式一对多自动匹配填写，支持列映射配置保存 |

## 使用说明

进入对应子目录安装依赖后运行：

```bash
pip install -r requirements.txt
```
