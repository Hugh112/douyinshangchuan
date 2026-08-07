# v3.0.0 构建验收记录

验收日期：2026-08-07

## 发布物

| 文件 | 大小（字节） | SHA256 |
| --- | ---: | --- |
| `release/DouyinPublisher_Setup_v3.0.0.exe` | 67073408 | `b8d19e751078cc665a7c99653c12055f7e3b0bbc1d97f328930a55b0adaf164f` |
| `packages/DouyinPublisher_Update_v3.0.0.zip` | 89822023 | `b5203682233d7d00e782ebd351d1c70568f31a89c693952ee5a55af0b58ea195` |

`version.json` 的版本、文件名、下载 URL、大小和 ZIP SHA256 已按上表生成，并随 v3.0.0 代码发布；统一管理后台版本策略需保持相同 URL 与哈希。

## 自动检查结果

- `py_compile`：`app_main.py`、`app_main.pyw`、`online_updater.py`、`platform_auth.py` 通过。
- 源码 `--self-test`：通过。
- 冻结版 EXE `--self-test`：通过。
- 自定义路径安装：通过；安装态 EXE 自检退出码 0；静默卸载后测试目录已清理。
- `app_main.py` 与 `app_main.pyw` SHA256 相同：`debc123abcc78e4d6b5a18a4d24a93502efd5fbd0e4d94cbf3ab4f0d5cb41ef2`。
- ZIP 顶层仅包含 `DouyinPublisher.exe`、`DouyinPublisherUpdater.exe`、`Start_Douyin_Publisher.vbs`、`_internal`。
- ZIP 无绝对路径、父目录穿越、`__pycache__`、`.pyc`、`.spec`、用户配置和历史版本启动器。
- 仓库内用户配置和 `%LOCALAPPDATA%` 安装版配置在构建前后 SHA256 未变化。

## 人工验收范围

自动测试不会向真实抖音账号发布内容。安装后请重点人工验证：标题输入框的当前平台 DOM、城市图片实际上传、同一共享图片池跨账号使用、发布成功后删除策略，以及真实发布后进入作品管理页的判断。
