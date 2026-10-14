"""Exercise the production publish functions with generated fixtures, without submitting."""
import argparse
import hashlib
import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "app"))
import app_main as app
from PIL import Image, ImageDraw


def snapshot_user_data():
    data_dir = Path(os.environ["LOCALAPPDATA"]) / "DouyinPublisher"
    paths = {
        app.CONFIG_PATH, data_dir / "douyin_gui_config.json", data_dir / "authorization.bin",
        data_dir / "login_preferences.bin", app.BITBROWSER_TOKEN_PATH,
        data_dir / "bitbrowser_connection_test.json",
    }
    for config_path in (app.CONFIG_PATH, data_dir / "douyin_gui_config.json"):
        if not config_path.is_file():
            continue
        cfg = json.loads(config_path.read_text(encoding="utf-8-sig"))
        paths.add(app.state_file_path(cfg))
        for account in app.browser_accounts_from_config(cfg):
            paths.add(app.state_file_path(app.config_for_browser_account(cfg, account)))
    return {path: hashlib.sha256(path.read_bytes()).hexdigest() if path.is_file() else None for path in paths}


def run():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", required=True)
    args = parser.parse_args()
    output = Path(args.output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    connection = json.loads((Path(os.environ["LOCALAPPDATA"]) / "DouyinPublisher" / "bitbrowser_connection_test.json").read_text(encoding="utf-8"))
    before = snapshot_user_data()
    account = app.normalize_browser_account({
        "id": "bitbrowser-live-test", "name": connection["window_name"], "browser_type": "bitbrowser",
        "bitbrowser_api_url": connection["api_url"], "bitbrowser_window_id": connection["window_id"],
        "browser_path": "", "use_schedule": True, "images_per_post": 2, "close_after_finish": False,
    }, 0, {})
    cfg = app.config_for_browser_account(app.default_config(), account)
    cfg.update({
        "state_path": str(output / "state.json"), "log_path": str(output / "log.csv"),
        "debug_dir": str(output / "screenshots"), "no_raise_browser": True,
        "close_chrome_before_start": False, "wait_min_seconds": 0.4, "wait_max_seconds": 0.7,
        "creator_page_ready_timeout_seconds": 120, "retry_times": 1, "overall_retry_times": 0,
        "upload_check_interval_seconds": 1, "upload_max_wait_seconds": 120,
        "declaration_enabled": True, "declaration_entry_text": "自主声明",
        "declaration_text": "无需添加自主声明",
        "delete_copy_after_success": False, "delete_image_after_success": False,
    })
    app.ACTIVE_CFG = cfg
    images = []
    for index, color in enumerate(("#e6f2ff", "#f1efe5"), start=1):
        image_path = output / f"local_api_test_{index}.jpg"
        picture = Image.new("RGB", (1440, 1920), color)
        draw = ImageDraw.Draw(picture)
        draw.text((80, 120), f"LOCAL API INTEGRATION TEST {index}", fill="#172033", font_size=48)
        draw.text((80, 210), "Publish page validation only", fill="#172033", font_size=36)
        picture.save(image_path, quality=92)
        images.append(image_path)
    title = "比特浏览器接口测试"
    text = "本地接口联调测试。此内容仅用于发布页面操作验证，不提交作品。 #测试 #知识科普"
    slot = (datetime.now() + timedelta(days=2)).replace(hour=14, minute=30, second=0, microsecond=0)
    report = {"window_name": account["name"], "steps": [], "submitted": False}
    manager = None
    page = None
    context = None
    workbench_pages = []
    api_calls = []
    original_client = app.bitbrowser_client

    def recording_client(api_url):
        client = original_client(api_url)
        request = client.request

        def recording_request(path, *args, **kwargs):
            api_calls.append(path)
            return request(path, *args, **kwargs)

        client.request = recording_request
        return client

    app.bitbrowser_client = recording_client
    try:
        info = app.open_browser_with_cdp(cfg)
        report["reused_existing_window"] = info["reused"]
        manager = app.sync_playwright().start()
        browser = manager.chromium.connect_over_cdp(app.browser_connection_endpoint(cfg, info))
        if not browser.contexts:
            raise RuntimeError("比特浏览器未提供已登录的默认上下文。")
        context = browser.contexts[0]
        workbench_pages = [p for p in context.pages if app.is_bitbrowser_workbench_page(p)]
        page = app.keep_single_browser_tab(context)
        if any(app.is_bitbrowser_workbench_page(p) for p in context.pages):
            raise RuntimeError("开始发布前默认工作台未关闭。")
        report["default_workbench_closed"] = all(p.is_closed() for p in workbench_pages)
        report["default_workbench_tabs_closed"] = len(workbench_pages)
        if len(context.pages) != 1:
            raise RuntimeError("比特窗口中未能只保留一个发布标签。")
        report["single_publish_tab"] = True
        if app.keep_single_browser_tab(context, page) is not page or len(context.pages) != 1:
            raise RuntimeError("重复整理时未复用原发布标签。")
        report["repeat_tab_setup_reused_same_page"] = True

        def step(name, action):
            app.wlog("本地实机验证：" + name)
            app.run_publish_step(cfg, page, name, action)
            report["steps"].append(name)

        step("进入图文发布页", lambda: app.enter_publish_page(cfg, page))
        step("上传两张独立测试图片", lambda: app.upload_images(cfg, page, images))
        step("等待图文编辑页", lambda: app.wait_after_upload_to_editor(cfg, page, 2))
        step("图片上传回读", lambda: (app.wait_uploaded_with_config(cfg, page, 2), app.validate_uploaded_image(cfg, page, 2)))
        step("填写并回读标题", lambda: (app.fill_title(cfg, page, title), app.validate_title_filled(cfg, page, title)))
        step("填写正文与选择话题", lambda: app.fill_copy(cfg, page, text))
        step("正文与话题最终回读", lambda: (app.validate_copy_filled(cfg, page, text), app.validate_topics_selected(cfg, page, text)))
        step("选择并校验发布声明", lambda: app.select_and_validate_declaration(cfg, page))
        step("选择并校验音乐", lambda: (app.select_music(cfg, page), app.validate_music_added(cfg, page)))
        step("定位发布设置", lambda: (app.close_music_panel_if_open(cfg, page), app.scroll_to_publish_settings(cfg, page), app.validate_publish_settings_visible(cfg, page)))
        step("设置定时发布", lambda: app.set_schedule(cfg, page, slot))
        if not app.verify_schedule_time(page, slot):
            raise RuntimeError("定时时间最终回读不一致。")
        step("发布前图片标题正文话题声明音乐校验", lambda: (
            app.validate_uploaded_image(cfg, page, 2), app.validate_title_filled(cfg, page, title),
            app.validate_copy_filled(cfg, page, text), app.validate_topics_selected(cfg, page, text),
            app.select_and_validate_declaration(cfg, page), app.validate_music_added(cfg, page),
        ))
        page.screenshot(path=str(output / "publish_before_submit.png"), full_page=True)
        report["result"] = "publish_preflight_passed"
        app.wlog("实机发布前校验通过；未调用提交函数或发布按钮。")
    except Exception as exc:
        report["result"] = "failed"
        report["error"] = str(exc)
        if page is not None:
            try:
                page.screenshot(path=str(output / "failure.png"), full_page=True)
            except Exception:
                pass
        raise
    finally:
        if page is not None:
            try:
                # 保留唯一窗口和标签；清除本次生成的测试内容，不关闭重开。
                page.goto(cfg["creator_url"], wait_until="domcontentloaded", timeout=60000)
            except Exception:
                pass
        if context is not None:
            report["final_publish_tab_count"] = len(context.pages)
        if manager is not None:
            manager.stop()
        changed = [path.name for path, expected in before.items()
                   if (hashlib.sha256(path.read_bytes()).hexdigest() if path.is_file() else None) != expected]
        report["changed_user_data_files"] = changed
        report["user_data_unchanged"] = not changed
        report["api_calls"] = api_calls
        report["browser_open_calls"] = api_calls.count("/browser/open")
        report["browser_close_calls"] = api_calls.count("/browser/close")
        app.bitbrowser_client = original_client
        (output / "result.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(report, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    run()
