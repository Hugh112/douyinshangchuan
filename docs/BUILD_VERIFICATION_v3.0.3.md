# v3.0.3 构建与验证记录

构建日期：2026-08-10
更新策略：可选更新，不强制升级

## 产物

- 安装包：`release/DouyinPublisher_Setup_v3.0.3.exe`
  - 大小：`67123522` 字节
  - SHA256：`10533d1104cfb8960c40cf3eece6095502140a726c146ed02001448e845d11d9`
- 在线更新包：`packages/DouyinPublisher_Update_v3.0.3.zip`
  - 大小：`89868738` 字节
  - SHA256：`d708efc6787c3c65dc8eb75373780e48dae117010722ef036ea3f22bdb3fc5b2`
- 主程序：`dist/DouyinPublisher/DouyinPublisher.exe`
  - 文件版本 / 产品版本：`3.0.3.0`
  - 大小：`10327688` 字节
  - SHA256：`00804be953bcfc5596157e37fb2ab47070ca577c06c9fe9b54f427898b64b1bc`
- 主程序源码：`app/app_main.py` / `app/app_main.pyw`
  - 两文件 SHA256：`22b26dac1551a35993905668f7dd956b7d4b30119a78774fb4cdd80ddbf710b3`

## 已执行验证

- `py_compile`：`app_main.py`、`app_main.pyw`、`online_updater.py`、`platform_auth.py` 通过。
- 源码与冻结版 `--self-test`：均返回 `SELF_TEST_OK`。
- `app_main.py` 与 `app_main.pyw`：逐字节及 SHA256 完全一致。
- 真实 Playwright 本地页面验证：原值 `2026-08-10 17:05` 可通过真实键盘改为目标 `2026-08-11 08:05`，输入框回读匹配。
- 发布页已到底部自检：滚轮调用次数 `0`、额外等待次数 `0`。
- GUI 布局测量：1240×880 窗口下任务控制可用宽度 `1184`、控件需求宽度 `1088`；重置进度与清空日志位于同一独立操作区。
- 更新 ZIP：共 `1874` 项，顶层仅 `_internal`、`DouyinPublisher.exe`、`DouyinPublisherUpdater.exe` 和固定名 `Start_Douyin_Publisher.vbs`。
- ZIP 路径检查：绝对路径、父目录穿越、`__pycache__`、`.pyc`、`.spec`、历史版本启动器、历史 README 和用户配置均为 `0` 项。
- 使用隔离 AppId 和仓库临时目录完成真实静默安装；安装后 EXE 自检通过，隔离卸载成功，测试目录及测试注册项均已移除。

GitHub Release 上传后还需下载附件并再次核对服务器返回文件的 SHA256 与本记录一致。
