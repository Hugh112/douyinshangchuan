# v3.1.2 包体验收记录

## 产物

- 安装包：`release/DouyinPublisher_Setup_v3.1.2.exe`
  - 大小：`77,094,958` 字节
  - SHA-256：`67D269BC07B3AEBF57DD611307BCC5A93733F37C03AA55E1498800C7500C6087`
- 更新包：`packages/DouyinPublisher_Update_v3.1.2.zip`
  - 大小：`100,647,872` 字节
  - SHA-256：`340B442ED016CE4177A55FACEFF286D9057D97DE9CBBAF93FEC5FD034CB58CCD`
- 安装后主程序 SHA-256：`3433FE6F797023F79423A7FDEFB7A162A03DAE65B7D01363E56160E4B5FAAEF3`
- `app_main.py` / `app_main.pyw` 共同 SHA-256：`12F2ACFFA2AD465DC016747F1A57D384986808C17E307B927913130B3BFC2ADC`

## 代码与实机回归

- Python 语法检查、主程序自检、更新器自检和授权单元测试通过；`pytest` 为 `4 passed`。
- 实机确认旧逻辑的点击前正文基准被旧 Locator 误读为空；网页实际正文在点击前含 `U+200B`，点击后正确变为“正文 + `#` + `U+200B`”。
- 修复后使用本机队列顺序 3“湖南2”执行完整发布前流程：`#岳阳县`、`#亲子鉴定`、`#亲子鉴定中心` 均点击完全同名候选并最终回读为 `3/3`。
- 网页最终写回内容包含 `U+00A0`、普通空格和 `U+200B`；清理不可见空字符后的正文和话题校验通过。
- 实机验证只到发布按钮之前，未点击最终发布，未消费正式 Excel 或发布进度。

## 包内容与安装

- 更新 ZIP 共 `1947` 个条目，根目录为 `DouyinPublisher.exe`、`DouyinPublisherUpdater.exe`、`Start_Douyin_Publisher.vbs` 和 `_internal`。
- ZIP 不含 `__pycache__`、`.pyc`、`.spec`、用户配置、授权、登录偏好、状态、日志或截图。
- 已覆盖安装到 `D:\软件\DouyinPublisher`，文件版本和产品版本均为 `3.1.2.0`。
- 安装后的主程序和更新器自检均返回 `0`，普通 GUI 启动存活检查通过后已关闭测试进程。
- 覆盖安装前后（首次启动前）`%LOCALAPPDATA%\DouyinPublisher` 的配置、DPAPI 授权和登录偏好 SHA-256 完全一致；随后执行正常 GUI 启动检查时，授权文件因会话刷新而更新，配置与登录偏好仍保持不变。

## 发布状态

- 根目录 `version.json` 已按最终 ZIP 的 URL、大小和 SHA-256 更新为 3.1.2。
- 尚未代替用户上传云服务器或登记统一管理后台。
