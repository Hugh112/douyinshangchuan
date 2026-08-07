# 在线更新与发布文档

## 1. 更新架构

客户端的清单优先级如下：

1. 统一管理后台：授权在线时调用产品 `publisher.douyin`、通道 `stable` 的发布策略。
2. GitHub Raw：`https://raw.githubusercontent.com/Hugh112/douyinshangchuan/main/version.json`
3. jsDelivr：GitHub Raw 的缓存灾备地址。

统一后台返回完整有效策略时拥有最高优先级。后台不可达、未配置版本，或声明了更高版本但缺少下载地址/合法 SHA256 时，客户端继续回退 GitHub。GitHub 清单请求会附加时间参数并发送 `no-cache` 头。

当前更新包由 GitHub Releases 托管。`version.json` 只保存版本策略与下载地址，不把 ZIP 提交进源码分支。

## 2. 更新执行

1. `available_update_info` 比较语义化数字版本。
2. 新版本必须提供 64 位十六进制 SHA256；否则拒绝更新。
3. `start_online_update` 启动 `DouyinPublisherUpdater.exe` 并退出主程序。
4. 更新器下载到临时目录并再次计算 SHA256。
5. 安全检查 ZIP 条目，拒绝绝对路径、父目录穿越和符号链接。
6. 覆盖固定文件名，保留清单声明的配置路径，并清理 `OBSOLETE_PATHS` 中明确列出的历史文件；固定名 `Start_Douyin_Publisher.vbs` 作为旧快捷方式到新 EXE 的兼容桥接，不得删除。
7. 优先无黑框启动固定 `DouyinPublisher.exe`，旧安装快捷方式继续通过固定 VBS 启动该 EXE，源码场景才回退 `.pyw`。

## 3. 发布新版本

1. 同步修改 `APP_VERSION`、窗口标题来源、`build_release.ps1`、Inno Setup 和 `version_info.txt`。
2. 保持 `app_main.py` 与 `app_main.pyw` 完全一致。
3. 执行语法检查、自检和核心回归。
4. 运行 `build_release.ps1` 生成安装包与更新 ZIP。
5. 检查 ZIP 根目录直接包含 `DouyinPublisher.exe`、`DouyinPublisherUpdater.exe`、`_internal` 和固定名兼容启动器，不得多套一层目录。
6. 确认 ZIP 不含 `__pycache__`、`.pyc`、`.spec`、版本号启动器、历史 README 或用户配置。
7. 计算最终 ZIP SHA256 和大小，写入根目录 `version.json`。
8. 在 GitHub 创建同版本 Release，上传文件名完全一致的更新 ZIP和安装包。
9. 先用 GitHub Raw 清单和 Release 下载地址做一次真实校验，再在统一后台登记相同版本、URL、SHA256、大小和发布说明。
10. 观察授权心跳、更新检查错误率和客户端遥测后再扩大投放比例。

## 4. `version.json` 必填字段

```json
{
  "latest_version": "3.0.0",
  "minimum_supported_version": "2.3.0",
  "download_url": "https://github.com/Hugh112/douyinshangchuan/releases/download/v3.0.0/DouyinPublisher_Update_v3.0.0.zip",
  "sha256": "最终 ZIP 的小写 SHA256",
  "package_size_bytes": 0,
  "force_update": false,
  "is_forced": false,
  "rollout_percentage": 100,
  "preserve_paths": ["app/douyin_gui_config.json"],
  "release_notes": ["更新说明"]
}
```

文件名、Release 标签、下载 URL、大小和 SHA256 必须来自同一个最终 ZIP。任何重打包都会改变哈希，需要重新写清单和后台记录。

同一版本号重新打包只用于人工测试或替换尚未投放的候选资产；已经安装相同版本号的客户端不会把它识别为升级。正式向已安装用户投放修订包时应提升补丁版本。

## 5. 回滚原则

客户端只执行升级，不自动降级。发现问题时不要把用户静默降回旧版：

- 立即在统一后台暂停投放或降低 rollout。
- 保留上一稳定版 Release 资产用于人工恢复。
- 修复后发布更高补丁版本，例如 3.0.1，并使用新的 SHA256。
- GitHub Raw 与统一后台应保持同一稳定目标；后台优先用于快速控制，GitHub 用于清单和包体灾备。

## 6. 用户数据保护

更新和安装不得覆盖或删除：

- `%LOCALAPPDATA%\DouyinPublisher\douyin_gui_config.json`
- DPAPI 授权与登录偏好
- 用户图片、Excel、日志、进度状态和调试截图

如需清理历史文件，只能在 `app/online_updater.py` 的 `OBSOLETE_PATHS` 中添加经过确认的相对路径，并在发布前检查没有匹配用户数据目录。
