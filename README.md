# 抖音单图文发布 - 在线更新仓库

当前版本：v2.2.3

本仓库用于提供软件在线更新文件。软件会读取根目录 `version.json`，然后下载 `packages` 目录中的更新包。

## v2.2.3 更新内容

- 修复上一版压缩包在部分 Windows 解压工具中中文文件名乱码的问题。
- 保留 v2.2.2 的 Excel/WPS 表格读取修复。
- 在线更新时继续保留本机配置：`app/douyin_gui_config.json`。

## 上传要求

请确保以下文件在仓库根目录：

- `version.json`
- `packages/douyin_single_publish_v2.2.3_encoding_fixed.zip`
- `README.md`

提交信息建议填写：

`fix zip filename encoding v2.2.3`
