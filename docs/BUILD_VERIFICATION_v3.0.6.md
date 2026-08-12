# v3.0.6 正式发布验收记录

日期：2026-08-12

本记录对应“音乐面板、目录型日志路径、行政区图片匹配及多浏览器 CDP”正式版本。v3.0.6 已在统一云服务器和统一管理后台按 Stable、100% 灰度、非强制策略发布。

## 候选产物

| 文件 | 大小（字节） | SHA256 |
| --- | ---: | --- |
| `release/DouyinPublisher_Setup_v3.0.6.exe` | 67223211 | `aab3f0e9ddd0042496f294e468b689296da49b504c88bea7ba1b646548bbfa58` |
| `packages/DouyinPublisher_Update_v3.0.6.zip` | 89971758 | `7eb90cd7df2987aacdb15f1c911261b4158be796d02e46c85826883e65a762da` |

## 验证结果

- `app_main.py`、`app_main.pyw`、`online_updater.py` 通过 `py_compile`。
- `app_main.py` 与 `app_main.pyw` 内容及 SHA256 完全一致。
- 源码主程序、冻结版主程序和冻结版更新器自检通过。
- 构造“普通空音乐模块 + music data 属性 + 绝对定位”异常页面，严格面板识别返回未打开，不再提前进入歌曲列表。
- 构造标准音乐弹窗，严格面板识别、面板根节点标记及歌曲行识别通过。
- 构造新版模块点击无效页面，程序在新版路径校验失败后进入完整 v3.0.1 识别流程，并通过音乐行右侧实际位置打开面板。
- 将状态/日志字段设置为现有目录时，分别解析为目录内固定名 `douyin_publish_state.json` 和 `douyin_publish_log.csv`；账号独立状态文件仍使用固定账号后缀。
- 构造日志路径不可写场景，日志写入返回失败但不抛出；已确认成功计数与状态文件在日志操作前落盘。
- 城市别名覆盖市、州、区、县、自治州、自治县、盟、旗等行政级别；广州等名称本身以“州”结尾的二字地名不会被错误截短；同一简称命中多个子目录时明确拒绝歧义匹配。
- `启动前关闭 Chrome 残留` 只清理 Chrome，不再结束 `node.exe`；账号 worker 和账号检测均在启动 Playwright 驱动前完成浏览器残留处理。
- 使用两个隔离用户目录和两个调试端口同时启动本机 Chrome，9222 逻辑对应的双 CDP 测试端口均成功监听并由 Playwright 连接；测试浏览器按精确 PID 关闭。
- 轮换模式下，缺少显式 `--user-data-dir` 的不同快捷方式/端口会得到不同且稳定的持久浏览器目录，避免 Chrome 合并进程后第二个调试端口不出现。
- ZIP 共 1874 项，根目录仅为 `_internal`、`DouyinPublisher.exe`、`DouyinPublisherUpdater.exe` 和固定名 `Start_Douyin_Publisher.vbs`。
- ZIP 不含绝对路径、父目录穿越、`__pycache__`、`.pyc`、`.spec`、版本号启动器、用户配置、发布状态、日志或调试截图。
- 安装包在仓库隔离目录完成自定义路径安装；安装后的主程序/更新器自检及 GUI 启动检查通过。
- 启动检查使用隔离 `LOCALAPPDATA`；真实 `douyin_gui_config.json` 的 SHA256 前后保持 `70366C7B00D14D3F723BE830130D1EBD69A1848383565701A9DFA3CC45E1F38C`。
- v3.0.4 未启用定时时直接提交的修复、v3.0.5 CDP 慢启动与运行状态刷新逻辑继续由内置自检覆盖。

## 正式发布状态

- 安装包固定下载地址：`https://api.xibao-zg.top/updates/publisher.douyin/stable/3.0.6/DouyinPublisher_Setup_v3.0.6.exe`；公网完整回下载大小和 SHA256 与上表一致。
- 更新 ZIP 固定下载地址：`https://api.xibao-zg.top/updates/publisher.douyin/stable/3.0.6/DouyinPublisher_Update_v3.0.6.zip`；公网完整回下载及 Range 请求通过，大小和 SHA256 与上表一致。
- 静态 Stable 清单与管理后台 `publisher.douyin` 均为版本代码 30006、最低支持 2.3.0、100% 灰度、非强制；v3.0.5 客户端可检测到 v3.0.6。
- 管理后台版本记录已发布；按用户决定，本版本产品公告不再发布。
- v3.0.6 包内更新器具备独立进度窗口。用户决定本版本不再重打包；从下一版本开始，发布门禁必须使用“上一线上稳定版真实安装客户端”完整发起更新并肉眼验证进度窗口，不能只自检新包内更新器。该门禁已写入 `docs/UPDATE_GUIDE.md` 和 `docs/HANDOVER.md`。

## 发布后人工观察项

- 在发生过误判的设备上确认日志先尝试新版模块点击；若模块无效，出现“切换到 v3.0.1 页面识别点击流程”，随后能够打开真实音乐抽屉。
- 在发生过 `Connection closed while reading from the driver` 的设备上，选择“启动前关闭 Chrome 残留”、不选择轮换发布，确认首个及后续账号均能连接各自端口。
- 对未显式配置用户目录的轮换账号，首次使用新隔离目录时完成一次抖音登录；显式用户目录的快捷方式应继续复用原登录状态。
- 继续观察音乐选择、立即发布和定时发布的真实账号运行情况；若发现回归，立即暂停后续通知并使用更高补丁版本修复，不对已发布的同版本包静默重打。
