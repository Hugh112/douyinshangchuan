# v3.0.7 正式版验收记录

发布日期：2026-09-03

## 产物

| 文件 | 大小（字节） | SHA-256 |
| --- | ---: | --- |
| `release/DouyinPublisher_Setup_v3.0.7.exe` | 77037918 | `a13d33e4e39be12cc196e4d4d5d274bf0659483f99f1a966ef0ea36cfdd90eb0` |
| `packages/DouyinPublisher_Update_v3.0.7.zip` | 100638123 | `64a4a5e9f2a6591e5d5cbceea200419fa292a0ca7d4ccfd3479e6894f548fbd4` |

## 已通过检查

- 用户已确认新版本异机测试可用，并授权正式推送。
- `app/app_main.py` 与 `app/app_main.pyw` SHA-256 完全一致。
- `py_compile`：主程序双入口、在线更新器、平台授权模块通过。
- 主程序 `.py` / `.pyw` 双入口 `--self-test` 通过。
- 更新器 `--self-test` 通过，版本为 3.0.7。
- `python -m unittest discover -s tests -v`：4 项授权登录回归通过。
- `git diff --check` 通过。
- 更新 ZIP 根目录直接包含 `DouyinPublisher.exe`、`DouyinPublisherUpdater.exe`、`Start_Douyin_Publisher.vbs` 和 `_internal`，无多余顶层目录。
- ZIP 未包含 `__pycache__`、`.pyc`、`.spec`、用户配置、发布状态、CSV 日志或调试截图。
- ZIP 内主程序 SHA-256 与构建目录主程序一致：`873d08b14c8e08a432d43a21b57c0b0525fa74e334975552a582771b74e1200c`。
- 已使用独立验证 AppId 静默安装到自定义路径；安装版 EXE `--self-test` 返回 0，文件/产品版本均为 3.0.7.0，随后正常卸载。
- 安装启动自检使用隔离 `LOCALAPPDATA`，未读取、删除或重置现有用户配置、授权文件、浏览器资料、图片、Excel、日志和发布进度。
- 3.0.6 正式安装包和更新 ZIP 原样保留，作为上一稳定回滚版本。
- 根据用户要求未控制桌面；更新进度通过更新器源码、打包内容及 `--self-test` 验证，正式版保留 v3.0.6 已加入的下载、校验、解压、安装和重启进度窗口。

## 发布后校验

- 2026-09-03 已在轻量应用服务器端校验人工上传的安装包和更新 ZIP：文件名、大小及 SHA-256 均与正式构建记录一致。
- 正式文件保存于数据盘 `/data/douyin-publisher/downloads/stable/3.0.7/`；现有项目级绑定将其发布到 `/srv/apps/ByxxPublisher/runtime/platform-production/updates/publisher.douyin/stable/3.0.7/`，未修改全局 Docker data-root，也未影响其他容器。
- 公网安装包完整下载后 SHA-256 为 `a13d33e4e39be12cc196e4d4d5d274bf0659483f99f1a966ef0ea36cfdd90eb0`。
- 公网更新 ZIP 完整下载后 SHA-256 为 `64a4a5e9f2a6591e5d5cbceea200419fa292a0ca7d4ccfd3479e6894f548fbd4`，Range 请求返回 `206 Partial Content` 和总长度 `100638123`。
- 原 v3.0.6 清单已备份到 `/data/douyin-publisher/backups/update-manifests/`，正式清单通过临时文件校验后原子替换。
- 公网清单已确认版本 3.0.7、最低兼容版本 2.3.0、非强制、灰度 100%，下载 URL、包大小和 SHA-256 均正确。
- 按用户要求，统一管理后台由用户自行登记：Stable、版本代码 30007、非强制、灰度 100%。
