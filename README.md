# 抖音单图文发布 - 在线更新仓库

当前版本：v2.2.4

本仓库用于提供软件在线更新文件。软件会读取根目录 `version.json`，然后下载 `packages` 目录中的更新包。

## v2.2.4 更新内容

- 修复文案文件误选 `.py` 脚本后，被 WPS/Excel 当成表格打开的问题。
- 文案文件选择框限制为 `.xlsx`、`.xls`、`.xlsm`、`.csv`、`.txt`，避免误选脚本文件。
- 保存配置和读取文案时增加文件格式校验，误选 `.py` 会直接弹窗提示。
- `read_copies` 增加 `.txt` / `.csv` 读取支持。
- 本次更新不包含强制绑定 Python 3.12 的改动，仍使用当前系统默认 `py`。

## 上传要求

请确保以下文件在仓库根目录：

- `version.json`
- `packages/douyin_single_publish_v2.2.4_wps_py_fix.zip`
- `README.md`

提交信息建议填写：

`fix wps py misselect v2.2.4`
