# v3.2.0 本机安装资料清理

日期：2026-10-14。最终包、两代保留包哈希、隔离安装/更新及用户数据保留均通过后执行。保留 3.2.0、3.1.3、3.1.1 三代安装包与更新 ZIP，另保留 3.2.0 独立更新器和校验文件。

删除前逐个解析并验证绝对路径，检查无重解析点、用户配置、凭据、用户图片、Excel 或进度文件。程序内置图标先与最终 ZIP 对应条目逐字节校验，再作为生成程序文件清理。未删除云服务器/GitHub 资产、Git 历史、正式安装目录、任何 runtime 或用户真实登录过的 stable-installed 目录。用户资料所在测试安装路径保留，只移除明确的生成程序文件。

本次释放 2071906085 字节（1.93 GiB）。清理对象为可重建安装/验包产物；需要时可由 Git 源码重新构建，但重新构建的文件哈希可能不同。三代保留包保持原文件与哈希，可用于人工安装恢复。

| 已删除绝对路径 | 字节数 |
| --- | ---: |
| D:\CodexProjects\douyin_single_publish\release\DouyinPublisher_Setup_v3.1.2.exe | 77094958 |
| D:\CodexProjects\douyin_single_publish\packages\DouyinPublisher_Update_v3.1.2.zip | 100647872 |
| D:\CodexProjects\douyin_single_publish\release_transition\rollback\v3.0.10-pre-topic-test-20261008 | 177732625 |
| D:\CodexProjects\douyin_single_publish\validation\v3.1.1-package-20261008-topic-test | 226088577 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-bitbrowser\DouyinPublisher_Validation_v3.2.0.exe | 77115995 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-bitbrowser\installed\_internal | 196605113 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-bitbrowser\installed\DouyinPublisher.exe | 17390836 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-bitbrowser\installed\DouyinPublisherUpdater.exe | 12112090 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-bitbrowser\installed\Start_Douyin_Publisher.vbs | 347 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-bitbrowser\installed\unins000.exe | 4410306 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-bitbrowser\installed\unins000.dat | 572971 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\Validation_v3.2.0.exe | 77134684 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\StableValidation_v3.1.1.exe | 77090580 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\StableVisualValidation_v3.1.1.exe | 77090575 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\running-stable-updater.exe | 12112216 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\visual-stable-updater.exe | 12112216 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\visual-replay-updater.exe | 12112216 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\stable-source | 226089492 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\__pycache__ | 11877 |
| D:\CodexProjects\douyin_single_publish\dist\DouyinPublisher | 226122900 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\new-installed\_internal | 196605113 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\new-installed\DouyinPublisher.exe | 17400897 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\new-installed\DouyinPublisherUpdater.exe | 12111960 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\new-installed\Start_Douyin_Publisher.vbs | 347 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\new-installed\unins000.exe | 4410306 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\new-installed\unins000.dat | 576969 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\stable-visual-installed\_internal | 196605113 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\stable-visual-installed\DouyinPublisher.exe | 17404216 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\stable-visual-installed\DouyinPublisherUpdater.exe | 12113224 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\stable-visual-installed\Start_Douyin_Publisher.vbs | 347 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\stable-visual-installed\unins000.exe | 4410306 |
| D:\CodexProjects\douyin_single_publish\validation\v3.2.0-release\stable-visual-installed\unins000.dat | 618841 |

本机机器可读记录：validation/v3.2.0-release/cleanup-result.json。历史发布说明保留；运行数据、校验日志与凭据不提交 Git。
