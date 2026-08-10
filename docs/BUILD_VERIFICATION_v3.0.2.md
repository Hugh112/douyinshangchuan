# v3.0.2 构建与验证记录

构建日期：2026-08-10
更新策略：可选更新，不强制升级

## 产物

- 安装包：`release/DouyinPublisher_Setup_v3.0.2.exe`
  - 大小：`67118005` 字节
  - SHA256：`e5c03b66cc2ad71ba41b2070abd01766072e44f2631ff3b33eab373cc71dff0a`
- 在线更新包：`packages/DouyinPublisher_Update_v3.0.2.zip`
  - 大小：`89864742` 字节
  - SHA256：`93da74acfe89bfdeb52f91521b06e8dc350c0cfb9b1d83680f035d50f597da33`
- 主程序源码：`app/app_main.py` / `app/app_main.pyw`
  - 两文件 SHA256：`400c7369a76e72a5d12291bf427f1663c420ff42d1758b535a90a0336974d411`

## 已执行验证

- `py_compile`：`app_main.py`、`app_main.pyw`、`online_updater.py`、`platform_auth.py`。
- 源码 `--self-test`：通过。
- `app_main.py` 与 `app_main.pyw` 内容及 SHA256：完全一致。
- 真实 CDP 页面只读探测：主音乐模块“未选择”和右侧 `sidesheet` 音乐面板均正确识别，推荐列表及“使用”按钮可读取。
- 活跃测试账号页面复检：音乐模块显示歌曲名、`00:16` 和“修改音乐”后，新代码直接跳过重复选择并返回成功；随后现有任务日志继续产生发布成功记录。为避免与仍在运行的旧版队列争用浏览器，未并行启动第二个完整发布 worker。
- 更新 ZIP：共 1874 项，顶层仅 `DouyinPublisher.exe`、`DouyinPublisherUpdater.exe`、固定名 `Start_Douyin_Publisher.vbs` 和 `_internal`；路径安全与禁入文件检查通过，未包含用户配置。
- 冻结版 EXE `--self-test`：返回 0，UTF-8 管道中包含 `SELF_TEST_OK`。
- 为避免打断当时仍在运行的旧版发布任务，使用相同 `dist` 载荷和隔离 AppId 执行自定义路径静默安装；安装返回 0、安装后 EXE 自检返回 0、卸载返回 0，测试目录已由卸载器移除。

最终 GitHub Release 上传后，应再次比对 Release 附件 SHA256 与本记录一致。
