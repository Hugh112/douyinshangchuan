# 抖音单图文发布 - 在线更新仓库

这是 `抖音单图文发布 v2.2.0` 的在线更新仓库。

## 仓库文件说明

- `version.json`：客户端启动时读取的版本配置文件。
- `packages/douyin_single_publish_v2.2.0_online_update.zip`：客户端实际下载的完整更新包。
- `source/douyin_single_publish_v2_2_online_update/`：当前版本源码与启动文件，方便后续维护。

## 客户端更新逻辑

程序启动后会读取：

`https://raw.githubusercontent.com/Hugh112/douyinshangchuan/main/version.json`

当 `version.json` 里的 `latest_version` 大于客户端内置版本号时，程序会弹窗提示用户是否更新。用户点“是”后，程序会下载 `download_url` 指向的 zip 包，校验 `sha256`，然后覆盖本机程序文件。

默认保留：

- `app/douyin_gui_config.json`

这样用户本机的浏览器路径、图片目录、Excel 路径、定时参数等配置不会被覆盖。

## 后续发布新版本怎么做

1. 修改 `source/douyin_single_publish_v2_2_online_update/` 里的程序文件。
2. 把整个程序目录重新压缩为：`packages/douyin_single_publish_v新版本号.zip`。
3. 计算 zip 的 SHA256。
4. 修改 `version.json`：
   - `latest_version` 改成新版本号，例如 `2.2.1`。
   - `download_url` 改成新 zip 的 raw 链接。
   - `sha256` 改成新 zip 的 SHA256。
   - `release_notes` 写本次更新内容。
5. 提交到 GitHub `main` 分支。
6. 用户打开旧版本后会自动提示更新，也可以手动点“检查更新”。

## 注意

已经发给用户的旧版如果没有在线更新功能，需要先让用户安装一次本仓库里的 `v2.2.0_online_update` 版本。安装过这版以后，后续就不用再重复发完整压缩包了，只需要更新 GitHub 仓库里的 `version.json` 和 `packages`。
