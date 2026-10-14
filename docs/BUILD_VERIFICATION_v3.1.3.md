# v3.1.3 包体验收记录

## 产物

- 安装包：`release/DouyinPublisher_Setup_v3.1.3.exe`
  - 大小：`77,103,491` 字节
  - SHA-256：`62ED1873752AE6500ED3698FBB5FF1E26F7E1CCEC64A50AB614C51D7D80E0E66`
- 更新包：`packages/DouyinPublisher_Update_v3.1.3.zip`
  - 大小：`100,649,702` 字节
  - SHA-256：`27FA1EE36CEA9D8B5566D7C446F9DDDE05A3A37360F00BC4393F3E6BE648707A`
- 安装后主程序 SHA-256：`BD7442C243078BFA474972E3437057300CA4FB142C840B184CABA411D0173553`
- `app_main.py` / `app_main.pyw` 共同 SHA-256：`451A4C3E5C7261EDB3F2C0DF3262CDBDEA1F6A339469C1DC257547E8BD00D803`

## 代码与实机回归

- Python 语法检查、主程序自检、更新器自检和授权单元测试通过；`pytest` 为 `4 passed`。
- 实机确认抖音输入长正文后会重建 `contenteditable` 节点；旧逻辑读取父容器文字并访问失效 Locator，分别造成真实正文被误排除和后续话题约 30 秒等待。
- 修复后使用本机队列顺序 3“湖南2”和 211 字正文验证 `#成都简阳市`、`#DNA鉴定`、`#热点`、`#知识科普`：4/4 搜索选择和最终回读通过。
- 从第一次点击“添加话题”到四个话题全部回读约 8 秒，第 2、3、4 个话题间隔约 2 秒。
- 实机验证只到发布按钮之前，未点击最终发布，未消费正式 Excel 或发布进度。

## 包内容与安装

- 更新 ZIP 共 `1947` 个条目，其中 `1923` 个文件条目；根目录为 `DouyinPublisher.exe`、`DouyinPublisherUpdater.exe`、`Start_Douyin_Publisher.vbs` 和 `_internal`。
- ZIP 不含 `__pycache__`、`.pyc`、`.spec`、用户配置、授权、登录偏好、状态、日志、截图或缓存。
- 冻结版主程序文件版本和产品版本均为 `3.1.3.0`；主程序和更新器自检均返回 `0`。
- 已覆盖安装到 `D:\软件\DouyinPublisher`，文件版本和产品版本均为 `3.1.3.0`。
- 安装后的主程序和更新器自检均返回 `0`；普通 GUI 启动 8 秒存活，窗口标题为“抖音智能发布中心 v3.1.3｜智能发布中心”，检查后只关闭本次测试进程。
- 覆盖安装前后（首次启动前）`%LOCALAPPDATA%\DouyinPublisher` 的配置、DPAPI 授权和登录偏好 SHA-256 完全一致。

## 发布状态

- 根目录 `version.json` 已按最终 ZIP 的 URL、大小和 SHA-256 更新为 3.1.3。
- 尚未代替用户上传云服务器或登记统一管理后台。
