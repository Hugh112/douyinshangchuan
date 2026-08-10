# v3.0.4 构建与验证记录

构建日期：2026-08-10
更新策略：可选更新，不强制升级

## 产物

- 安装包：`release/DouyinPublisher_Setup_v3.0.4.exe`
  - 大小：`67121788` 字节
  - SHA256：`4dae40412dbc62119b06d4ae845c73d7c220b2e21f9ae087ac56bff4e25a6c4f`
- 在线更新包：`packages/DouyinPublisher_Update_v3.0.4.zip`
  - 大小：`89869975` 字节
  - SHA256：`be8a3b2c0e745ff826bd06876f1194f06bbf64e51a762144acf8a946005f4b21`
- 主程序：`dist/DouyinPublisher/DouyinPublisher.exe`
  - 文件版本：`3.0.4.0`
  - 大小：`10330044` 字节
  - SHA256：`77306e8350c67cdd58890139b02ee52c3ae3dbbf17b830ce2602052aeb1a5aff`
- 主程序源码：`app/app_main.py` / `app/app_main.pyw`
  - 两文件 SHA256：`383ddf75d5a96c8c4296486e9abe157cdc11c27ae6751452aef4aad014a5c9ae`

## 验证结果

- `py_compile`、源码 `--self-test`、冻结版 `--self-test`：通过。
- `app_main.py` 与 `app_main.pyw`：逐字节及 SHA256 完全一致。
- 真实抖音页面首次验证：成功识别无 `checked/aria-checked` 的新版定时模式，将旧值写为 `2026-08-11 08:17` 并回读一致。
- 真实抖音页面清空验证：`2026-08-11 08:17 → 空值 → 2026-08-11 08:36`，最终回读一致。
- v3.0.4 最终顺序验证：`2026-08-11 08:36 → 空值 → 2026-08-11 08:54`，直接使用 v3.0.1 DOM 稳定逻辑，最终回读一致。
- 真实测试仅上传测试图片并修改定时时间，没有点击底部发布，没有消费图片、文案或任务进度。
- 更新 ZIP 共 `1874` 项，顶层仅 `_internal`、`DouyinPublisher.exe`、`DouyinPublisherUpdater.exe` 和固定名 `Start_Douyin_Publisher.vbs`。
- ZIP 中绝对路径、父目录穿越、`__pycache__`、`.pyc`、`.spec`、历史启动器、历史 README 和用户配置均为 `0` 项。
