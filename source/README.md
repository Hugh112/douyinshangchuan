# 抖音单图文发布工具

当前最新版本：v2.2.1

## 在线更新

软件启动后会读取仓库根目录的 `version.json`，发现新版本后提示用户更新。

## v2.2.1 修复

- 修复文案Excel误选为 `.zip` 压缩包时出现 `io.excel.zip.reader` 报错的问题。
- 文案Excel选择窗口增加 Excel 文件筛选。
- 读取文案前增加中文格式校验，选错文件会提示重新选择 `.xlsx / .xls / .xlsm`。

## 上传更新方法

把本包解压后的 `version.json` 和 `packages/douyin_single_publish_v2.2.1_online_update_fixed.zip` 上传到 GitHub 仓库根目录对应位置即可。
