# 在线更新与发布文档

## 1. 更新架构

客户端的清单优先级如下：

1. 统一管理后台：授权在线时调用产品 `publisher.douyin`、通道 `stable` 的发布策略。
2. 统一云服务器静态清单：`https://api.xibao-zg.top/updates/publisher.douyin/stable/manifest.json`

统一后台返回完整有效策略时拥有最高优先级。后台不可达、未配置版本，或声明了更高版本但缺少下载地址/合法 SHA256 时，客户端继续回退云服务器静态清单。清单请求会附加时间参数并发送 `no-cache` 头。

更新包由统一云服务器 `/updates/publisher.douyin/` 托管。根目录 `version.json` 是同结构的源码备份，不把 ZIP 或安装包提交进源码分支；GitHub 不再是客户端包体下载源。

## 2. 更新执行

1. `available_update_info` 比较语义化数字版本。
2. 新版本必须提供 64 位十六进制 SHA256；否则拒绝更新。
3. `start_online_update` 启动 `DouyinPublisherUpdater.exe` 并退出主程序。
4. 更新器立即显示独立进度窗口；下载阶段显示百分比、已下载/总大小和重试次数。
5. SHA256、ZIP 路径安全检查、解压、等待主程序退出、覆盖文件、历史清理和重启都显示对应阶段及总进度。
6. 安全检查 ZIP 条目，拒绝绝对路径、父目录穿越和符号链接。
7. 覆盖固定文件名，保留清单声明的配置路径，并清理 `OBSOLETE_PATHS` 中明确列出的历史文件；固定名 `Start_Douyin_Publisher.vbs` 作为旧快捷方式到新 EXE 的兼容桥接，不得删除。
8. 优先无黑框启动固定 `DouyinPublisher.exe`，旧安装快捷方式继续通过固定 VBS 启动该 EXE，源码场景才回退 `.pyw`。

## 3. 发布新版本

1. 同步修改 `APP_VERSION`、窗口标题来源、`build_release.ps1`、Inno Setup 和 `version_info.txt`。
2. 保持 `app_main.py` 与 `app_main.pyw` 完全一致。
3. 执行语法检查、自检和核心回归。
4. 运行 `build_release.ps1` 生成安装包与更新 ZIP。
5. 在隔离环境安装上一线上稳定版，从该真实安装目录发起一次到本候选的完整更新；必须肉眼确认独立进度窗口立即出现，并持续显示下载百分比/大小、SHA256、ZIP 检查、解压、安装和重启阶段。只验证新包内 `DouyinPublisherUpdater.exe` 不算通过，因为升级过程首先调用的是上一版已安装的更新器。任何阶段无可见进度都必须停止发布并提升版本修复，不得把“更新完成后新更新器具备进度窗口”当作本次跨版本验证。
6. 检查 ZIP 根目录直接包含 `DouyinPublisher.exe`、`DouyinPublisherUpdater.exe`、`_internal` 和固定名兼容启动器，不得多套一层目录。
7. 确认 ZIP 不含 `__pycache__`、`.pyc`、`.spec`、版本号启动器、历史 README 或用户配置。
8. 计算最终 ZIP SHA256 和大小，写入根目录 `version.json`。
9. 把安装包、更新 ZIP 和 `manifest.json` 上传到统一云服务器固定产品/通道目录；同一文件先以临时名上传，服务端校验 SHA256 和大小后再原子替换。
10. 通过公网 HTTPS 对清单、完整下载、Range 请求、大小和 SHA256 做一次真实校验，再在统一后台登记相同版本、云服务器 URL、SHA256、大小和发布说明。
11. GitHub 只推送已验证源码、版本清单和发布说明，可创建同版本标签作为源码备份，但不让客户端读取 GitHub 包体。
12. 观察登录验证、更新检查错误率和客户端遥测后再扩大投放比例；默认 `login_only` 客户端不得出现周期性授权心跳或隐式刷新。

## 4. `version.json` 必填字段

```json
{
  "latest_version": "3.0.6",
  "minimum_supported_version": "2.3.0",
  "download_url": "https://api.xibao-zg.top/updates/publisher.douyin/stable/3.0.6/DouyinPublisher_Update_v3.0.6.zip",
  "sha256": "最终 ZIP 的小写 SHA256",
  "package_size_bytes": 0,
  "force_update": false,
  "is_forced": false,
  "rollout_percentage": 100,
  "preserve_paths": ["app/douyin_gui_config.json"],
  "release_notes": ["更新说明"]
}
```

文件名、云服务器下载 URL、大小和 SHA256 必须来自同一个最终 ZIP。任何重打包都会改变哈希，需要重新写静态清单、源码清单和后台记录。

同一版本号重新打包只用于人工测试或替换尚未投放的候选资产；已经安装相同版本号的客户端不会把它识别为升级。正式向已安装用户投放修订包时应提升补丁版本。

## 5. 回滚原则

客户端只执行升级，不自动降级。发现问题时不要把用户静默降回旧版：

- 立即在统一后台暂停投放或降低 rollout。
- 保留上一稳定版云服务器资产用于人工恢复。
- 修复后发布更高补丁版本，例如 3.0.1，并使用新的 SHA256。
- 云服务器静态清单与统一后台应保持同一稳定目标；后台优先用于快速控制，静态清单用于授权后台暂时不可达时回退。

### 5.1 本机回滚与构建资料保留

- 本机每个正式发布范围保留“当前已验证安装包/更新包 + 最近一代已验证稳定回滚包”。新回滚代完成签名或 SHA-256、包内容、安装启动和用户数据保留验证后，可清理更旧的本机重复安装包、更新 ZIP、验包解压目录、临时构建产物和缓存。
- 云服务器更新资产和 GitHub 源码标签不属于本机磁盘清理对象。被静态清单、`version.json`、统一后台、正在执行的灰度投放或最新回滚说明引用的产物不得删除。
- 清理永远排除 `%LOCALAPPDATA%\DouyinPublisher`、用户图片/Excel、安装路径、进度状态、密钥、Git 历史、当前有效签名产物和最新已验证回滚包。删除前必须核对绝对路径与引用关系，删除后记录路径、大小和可恢复性。

## 6. 用户数据保护

更新和安装不得覆盖或删除：

- `%LOCALAPPDATA%\DouyinPublisher\douyin_gui_config.json`
- DPAPI 授权与登录偏好
- 用户图片、Excel、日志、进度状态和调试截图

如需清理历史文件，只能在 `app/online_updater.py` 的 `OBSOLETE_PATHS` 中添加经过确认的相对路径，并在发布前检查没有匹配用户数据目录。
