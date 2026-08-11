# v3.0.5 构建与发布验证

验证日期：2026-08-11

## 最终产物

| 产物 | 大小 | SHA256 |
| --- | ---: | --- |
| `DouyinPublisher_Update_v3.0.5.zip` | 89870261 | `d2a0ebd23fd9ca466abc485850d052be43fa9f682dd9ee6cca1852f61fd82c67` |
| `DouyinPublisher_Setup_v3.0.5.exe` | 67131740 | `cc4b7e66ab5eda0f2965ebf6d16cba5e01b18a01b9896732319c265a475dac70` |

## 验证结果

- `py_compile`：`app_main.py`、`app_main.pyw`、`online_updater.py`、`platform_auth.py` 通过。
- 内置 `--self-test`：通过。
- `app_main.py` 与 `app_main.pyw`：逐字节 SHA256 一致。
- 音乐 DOM 回归：加载中抽屉识别为已打开；列表就绪后歌曲行和“使用”按钮均可识别。
- 更新 ZIP：1874 个条目；根层仅 `_internal`、两个固定 EXE 和 `Start_Douyin_Publisher.vbs`。
- ZIP 禁入项：无绝对路径、`..`、`__pycache__`、`.pyc`、`.spec`、版本号启动器或历史 README。
- 自定义路径安装：安装器退出码 0；固定 VBS 指向 `DouyinPublisher.exe`。
- 安装版 EXE 自检：退出码 0；文件版本和产品版本均为 `3.0.5.0`。
- 用户数据保护：安装和安装版自检前后，本机现有授权/登录偏好文件哈希保持一致。
- 云服务器：清单、ZIP 和安装包均返回 HTTP 200；ZIP 支持 HTTP 206 Range。
- 公网完整重下：大小 89870261，SHA256 与本地最终 ZIP 完全一致。
- 统一后台端到端：测试授权账号读取到 3.0.5、云服务器 URL、正确 SHA256/大小、非强制、100% 投放及 v3.0.5 公告。

没有在本机构造真实抖音作品并点击最终发布，以免额外消费用户当前图片、文案和账号配额；音乐流程使用隔离网页 DOM 回归，端口修复使用代码路径和安装版自检验证。受影响电脑仍应进行一次人工账号发布确认。
