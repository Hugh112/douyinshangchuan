# 抖音智能发布中心 v3.0.7 更新公告

发布日期：2026-09-03

更新策略：Stable，非强制更新，灰度 100%，最低支持版本 2.3.0。

## 本轮更新

- 修复音乐入口偶发“点击动作执行了，但面板实际没有打开”的问题；新版模块操作无效时自动使用 v3.0.1 页面识别与真实点击兜底。
- 网页崩溃时自动刷新并安全重跑当前作品，不提前消费图片、文案或发布进度。
- 修复轮换发布共用 Excel 时文案重复、表格仍有文案却误报用完、补充文案后序号超限。
- 自动启动建立新的账号配额与排期轮次，同时保留未消费文案游标；手动停止和普通重启继续恢复当前进度。
- 城市图片优先匹配文案同名目录，再匹配市、州、区、县及“地级市 + 区县”层级别名；大量子目录使用缓存索引。
- 发布失败时采集抖音页面的封禁、限制、违规或字段错误原文，并写入账号状态、日志和失败提醒。
- 保留 v3.0.4 立即发布热补丁、v3.0.5 CDP 慢启动与状态实时刷新、v3.0.6 多浏览器隔离和在线更新进度显示。

## 正式产物

- 安装包：`DouyinPublisher_Setup_v3.0.7.exe`
- 安装包大小：`77037918` 字节
- 安装包 SHA-256：`a13d33e4e39be12cc196e4d4d5d274bf0659483f99f1a966ef0ea36cfdd90eb0`
- 更新包：`DouyinPublisher_Update_v3.0.7.zip`
- 更新包大小：`100638123` 字节
- 更新包 SHA-256：`64a4a5e9f2a6591e5d5cbceea200419fa292a0ca7d4ccfd3479e6894f548fbd4`

## 下载地址

- 安装包：`https://api.xibao-zg.top/updates/publisher.douyin/stable/3.0.7/DouyinPublisher_Setup_v3.0.7.exe`
- 更新包：`https://api.xibao-zg.top/updates/publisher.douyin/stable/3.0.7/DouyinPublisher_Update_v3.0.7.zip`
- 版本清单：`https://api.xibao-zg.top/updates/publisher.douyin/stable/manifest.json`
