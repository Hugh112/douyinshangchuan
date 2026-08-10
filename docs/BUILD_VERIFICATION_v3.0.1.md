# v3.0.1 构建验收记录

构建日期：2026-08-07

## 产物

| 文件 | 字节数 | SHA256 |
|---|---:|---|
| `release/DouyinPublisher_Setup_v3.0.1.exe` | 67074518 | `540b8e85b78e91b7380a63f466111fdd2bba20b887cbf48adb9fea3e17e0dad5` |
| `packages/DouyinPublisher_Update_v3.0.1.zip` | 89822872 | `d7971a0cbd2651c63b22830c275f2190584562c1f00660c60f25e2deae32a7fe` |

## 验收结果

- `app_main.py`、`app_main.pyw`、`online_updater.py` 通过 `py_compile`。
- `app_main.py` 与 `app_main.pyw` 内容及 SHA256 完全一致。
- 内置自测通过；打包后的 `DouyinPublisher.exe --self-test` 退出码为 0。
- 示例表格通过实际程序校验，4 行分别匹配济南、青岛、烟台、德州子文件夹。
- 更新 ZIP 共 1874 个条目，无绝对路径；顶层仅 `_internal`、主程序、更新器和固定名称启动器。
- 更新 ZIP 不含 `__pycache__`、`pyc`、`spec`、历史版本启动器、历史 README 或用户配置文件。
- 安装器保留自定义安装路径页面（`DisableDirPage=no`）。
- `version.json` 的版本、ZIP 文件名、大小与 SHA256 已按本地候选产物生成。

## 发布状态

本次仅生成本地候选产物，尚未推送 GitHub Release、线上 `version.json` 或统一管理后台。
