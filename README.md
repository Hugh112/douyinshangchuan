# 抖音智能发布中心

当前稳定版：`v2.5.1`

Windows 桌面端抖音图文自动发布工具。当前版本已接入统一授权后台，支持多浏览器账号按队列发布、账号专属内容、定时启动、运行遥测和安全在线更新。视频发布与 AI 内容生成属于后续版本范围，本版不包含这两项。

## 当前能力

- 软件账号登录、退出、切换账号、记住账号和 DPAPI 加密记住密码。
- 定时授权心跳；授权撤销或任务自动暂停时即时提醒。
- 多浏览器账号独立配置、发布配额、账号专属图片与 Excel、掉号跳过和顺序切换。
- 随机图片或按文件名排序取第一张图片。
- 仅允许读取 `.xlsx`、`.xls`、`.xlsm` 文案文件。
- 正文和话题使用浏览器键盘事件逐字输入，不占用剪贴板。
- 立即发布、抖音定时发布、多日排期、按星期和时间自动启动。
- 失败按前端配置重试；成功后按配置删除已用图片和文案。
- 暂停、继续、停止、重置进度和全局可见运行日志。
- 原生 `DouyinPublisher.exe` 无黑框启动，安装包内置运行依赖并支持自定义安装路径。

## 正式代码

- 主程序：`app/app_main.py`
- 同步入口：`app/app_main.pyw`，内容必须与主程序完全一致
- 授权客户端：`app/platform_auth.py`
- 在线更新器：`app/online_updater.py`
- 用户配置：`app/douyin_gui_config.json`，仅保留在用户电脑，不提交 Git
- 构建脚本：`build_release.ps1`
- 安装器定义：`installer/DouyinPublisher.iss`
- 更新清单：`version.json`

安装版把配置、授权会话和登录偏好存放在 `%LOCALAPPDATA%\DouyinPublisher`，不放入安装目录，因此升级或更换安装路径不会覆盖这些数据。

## 更新来源

客户端按以下顺序检查：

1. 已授权时优先读取统一管理后台 `publisher.douyin` 的稳定通道策略。
2. 后台不可达、未配置版本或返回的新版本策略不完整时，回退 GitHub Raw 根目录 `version.json`。
3. GitHub Raw 不可用时，使用 jsDelivr 备用清单。

当前更新包下载地址使用 GitHub Releases。所有更新都必须通过 SHA256 校验，更新器还会检查 ZIP 路径安全，并保留用户配置后再无黑框重启。

## 构建

在 Windows PowerShell 中运行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build_release.ps1
```

构建结果：

- `release/DouyinPublisher_Setup_v2.5.1.exe`
- `packages/DouyinPublisher_Update_v2.5.1.zip`

构建脚本会把依赖冻结到安装目录，并在成功后清理临时 `build`、`.spec`、`__pycache__` 和 `.pyc`。

详细维护信息见 [交接文档](docs/HANDOVER.md) 和 [更新发布文档](docs/UPDATE_GUIDE.md)。
