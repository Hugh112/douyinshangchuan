# -*- coding: utf-8 -*-
"""
抖音智能发布中心 v3.0.4 - 定时时间输入修复版
功能：
1. GUI 前端配置浏览器、图片、Excel、定时、等待、上传检测、重试等。
2. 自动打开抖音创作者平台图文发布页。
3. 上传图片后，以"编辑图片 / 已添加1张图片 / 封面设置 / 预览图文"等判断上传完成。
4. 填写文案和话题。
5. 只使用平台音乐面板，悬停音乐后点击"使用"。
6. 音乐成功后直接下滑到发布设置。
7. 支持定时发布 / 立即发布开关。
8. 点击发布前再次校验图片、文案、音乐、发布设置。
9. 支持随机使用图片（运行选项中勾选即可）。
10. 支持多个浏览器账号按配额顺序发布，也可勾选轮换发布并自定义“每个账号连续发布多少条后切换”。
11. 支持按星期和时间自动启动任务。
12. 保留在线更新、无黑框普通启动和完整调试启动。
13. 支持 Excel 城市、标题、文案三列绑定；兼容旧无表头单列文案。
"""

import copy
import csv
import hashlib
import json
import math
import os
import queue
import random
import re
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse
import urllib.request
import ctypes
import shutil
import uuid
from datetime import datetime, timedelta
from pathlib import Path

def _windowed_standard_stream(fd):
    """仅在 Windows 标准句柄真实有效时恢复管道；普通无黑框启动使用 NUL。"""
    duplicate = None
    try:
        if os.name == "nt":
            import msvcrt
            handle = msvcrt.get_osfhandle(fd)
            if handle in (-1, 0):
                raise OSError("invalid Windows standard handle")
            ctypes.set_last_error(0)
            file_type = ctypes.windll.kernel32.GetFileType(ctypes.c_void_p(handle))
            if file_type == 0 and ctypes.get_last_error() != 0:
                raise OSError("unavailable Windows standard handle")
        duplicate = os.dup(fd)
        return os.fdopen(duplicate, "w", encoding="utf-8", errors="replace", buffering=1)
    except Exception:
        if duplicate is not None:
            try:
                os.close(duplicate)
            except Exception:
                pass
        return open(os.devnull, "w", encoding="utf-8")


if sys.stdout is None:
    sys.stdout = _windowed_standard_stream(1)
if sys.stderr is None:
    sys.stderr = _windowed_standard_stream(2)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import pandas as pd
from openpyxl import Workbook, load_workbook

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
except Exception:
    tk = None

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

from platform_auth import (
    DpapiTokenStore,
    PlatformAuthClient,
    PlatformAuthError,
    STATE_AUTHENTICATING,
    STATE_AUTHORIZED,
    STATE_FORCE_LOGOUT,
    STATE_GRACE,
    STATE_SIGNED_OUT,
)

IS_FROZEN = bool(getattr(sys, "frozen", False))
SOURCE_APP_DIR = Path(__file__).resolve().parent
INSTALL_DIR = Path(sys.executable).resolve().parent if IS_FROZEN else SOURCE_APP_DIR.parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", SOURCE_APP_DIR))
APP_DIR = INSTALL_DIR if IS_FROZEN else SOURCE_APP_DIR
APP_DATA_DIR = (
    Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "DouyinPublisher"
    if IS_FROZEN
    else SOURCE_APP_DIR
)
CONFIG_PATH = APP_DATA_DIR / "douyin_gui_config.json"
AUTH_TOKEN_PATH = APP_DATA_DIR / "authorization.bin"
AUTH_LOGIN_PREFERENCES_PATH = APP_DATA_DIR / "login_preferences.bin"
DEBUG_SCREENSHOT_CLEANUP_STATE_PATH = APP_DATA_DIR / "debug_screenshot_cleanup.json"
APP_VERSION = "3.0.4"
APP_NAME = f"抖音智能发布中心 v{APP_VERSION}"
LOGO_ICO = RESOURCE_DIR / "assets" / "app_logo.ico"
LOGO_PNG = RESOURCE_DIR / "assets" / "app_logo.png"
ACTIVE_CFG = None
COUNTDOWN_EVENT_PREFIX = "__DOUYIN_COUNTDOWN__:"
BROWSER_STATUS_EVENT_PREFIX = "__DOUYIN_BROWSER_STATUS__:"
AUTO_PAUSE_EVENT_PREFIX = "__DOUYIN_AUTO_PAUSE__:"

AUTH_API_BASE_URL = "https://api.xibao-zg.top"
AUTH_PRODUCT_CODE = "publisher.douyin"
AUTH_RELEASE_CHANNEL = "stable"
AUTH_HEARTBEAT_SECONDS = 60

# 更新顺序：已授权时优先读取统一后台；后台不可达或未返回有效策略时，
# 回退 GitHub raw，并以 jsDelivr 作为最后备用。HTTP 清单请求均追加防缓存参数。
UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/Hugh112/douyinshangchuan/main/version.json"
UPDATE_MANIFEST_FALLBACK_URLS = [
    "https://cdn.jsdelivr.net/gh/Hugh112/douyinshangchuan@main/version.json",
]
UPDATE_TIMEOUT_SECONDS = 12
PRESERVE_UPDATE_PATHS = [
    "app/douyin_gui_config.json",  # 兼容旧版首次迁移
]


def parse_version(v):
    nums = []
    for part in re.findall(r"\d+", str(v or "0"))[:4]:
        try:
            nums.append(int(part))
        except Exception:
            nums.append(0)
    while len(nums) < 4:
        nums.append(0)
    return tuple(nums)


def urlopen_text_with_retry(url, timeout=UPDATE_TIMEOUT_SECONDS, retry=3):
    """读取更新配置，带重试和防缓存，解决 GitHub/CDN 偶发 SSL EOF 或缓存旧版本。"""
    last_err = None
    sep = "&" if "?" in url else "?"
    url_with_cache_buster = f"{url}{sep}_t={int(time.time())}"
    headers = {
        "User-Agent": f"DouyinPublisher/{APP_VERSION}",
        "Accept": "application/json,text/plain,*/*",
        "Cache-Control": "no-cache",
        "Pragma": "no-cache",
        "Connection": "close",
    }
    for i in range(max(1, int(retry))):
        try:
            req = urllib.request.Request(url_with_cache_buster, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read().decode("utf-8-sig", errors="replace")
        except Exception as e:
            last_err = e
            time.sleep(1.5 + i)
    raise last_err


def platform_release_to_manifest(policy):
    """把统一后台的 release-policy 映射成旧更新器兼容的 version.json 结构。"""
    policy = policy if isinstance(policy, dict) else {}
    version = policy.get("version") if isinstance(policy.get("version"), dict) else {}
    latest = str(version.get("latestVersion") or "").strip()
    return {
        "latest_version": latest,
        "minimum_supported_version": str(version.get("minimumSupportedVersion") or ""),
        "download_url": str(version.get("downloadUri") or ""),
        "sha256": str(version.get("fileSha256") or "").lower(),
        "package_size_bytes": version.get("packageSizeBytes"),
        "release_notes": str(version.get("releaseNotes") or ""),
        "is_forced": bool(version.get("isForced") or version.get("updateRequired")),
        "rollout_percentage": int(version.get("rolloutPercentage") or 100),
        "announcements": policy.get("announcements") or [],
        "source": "unified-platform",
    }


def fetch_update_manifest(auth_client=None):
    """统一后台优先；未登录或后台不可达时回退 GitHub raw，再回退 jsDelivr。"""
    errors = []
    if auth_client is not None and auth_client.state == STATE_AUTHORIZED:
        try:
            manifest = platform_release_to_manifest(auth_client.get_release_policy(timeout=6))
            if not isinstance(manifest, dict):
                raise ValueError("统一后台版本策略格式不正确")
            latest = str(manifest.get("latest_version") or "").strip()
            if not latest:
                raise ValueError("统一后台尚未配置可用版本")
            if parse_version(latest) > parse_version(APP_VERSION):
                download_url = str(manifest.get("download_url") or "").strip()
                sha256 = str(manifest.get("sha256") or "").strip().lower()
                if not download_url or not re.fullmatch(r"[0-9a-f]{64}", sha256):
                    raise ValueError("统一后台新版本策略缺少下载地址或合法 SHA256")
            return manifest
        except Exception as exc:
            errors.append(f"统一后台：{exc}")

    urls = [UPDATE_MANIFEST_URL] + list(UPDATE_MANIFEST_FALLBACK_URLS)
    last_err = None
    for url in urls:
        try:
            data = urlopen_text_with_retry(url)
            manifest = json.loads(data)
            if not isinstance(manifest, dict):
                raise ValueError("版本配置文件格式不正确")
            manifest["source"] = "github-raw" if "raw.githubusercontent.com" in url else "jsdelivr"
            return manifest
        except Exception as e:
            last_err = e
            errors.append(f"{url}：{e}")
            continue
    raise RuntimeError("；".join(errors) or str(last_err or "检查更新失败"))


def available_update_info(auth_client=None):
    manifest = fetch_update_manifest(auth_client)
    latest = str(manifest.get("latest_version") or manifest.get("version") or "").strip()
    download_url = str(manifest.get("download_url") or "").strip()
    if not latest:
        raise ValueError("version.json 缺少 latest_version")
    if parse_version(latest) <= parse_version(APP_VERSION):
        return None
    if not download_url:
        raise ValueError("version.json 缺少 download_url")
    sha256 = str(manifest.get("sha256") or "").strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", sha256):
        raise ValueError("更新策略缺少合法的 SHA256，已拒绝下载")
    return manifest


def format_release_notes(manifest):
    notes = manifest.get("release_notes") or manifest.get("notes") or ""
    if isinstance(notes, list):
        notes = "\n".join("- " + str(x) for x in notes)
    return str(notes).strip()


BROWSER_QUEUE_COLUMNS = (
    "publish", "order", "name", "path", "excel", "schedule", "port", "quota", "published", "status",
)
BROWSER_QUEUE_COLUMN_HEADINGS = {
    "publish": "发布", "order": "顺序", "name": "账号名称", "path": "浏览器快捷方式 / EXE",
    "excel": "文案 Excel", "schedule": "独立发布时间",
    "port": "端口", "quota": "发布条数", "published": "已发布", "status": "运行状态",
}
BROWSER_QUEUE_DEFAULT_WIDTHS = {
    "publish": 58, "order": 48, "name": 105, "path": 250,
    "excel": 190, "schedule": 230, "port": 66, "quota": 82, "published": 76, "status": 155,
}
BROWSER_QUEUE_MIN_WIDTHS = {
    "publish": 46, "order": 42, "name": 72, "path": 120,
    "excel": 100, "schedule": 130, "port": 52, "quota": 68, "published": 68, "status": 100,
}


def normalize_browser_queue_column_widths(raw):
    """只接受已知列和合理像素范围，避免损坏配置让表格不可用。"""
    raw = raw if isinstance(raw, dict) else {}
    result = {}
    for key in BROWSER_QUEUE_COLUMNS:
        try:
            width = int(raw.get(key, BROWSER_QUEUE_DEFAULT_WIDTHS[key]))
        except Exception:
            width = BROWSER_QUEUE_DEFAULT_WIDTHS[key]
        result[key] = max(BROWSER_QUEUE_MIN_WIDTHS[key], min(1200, width))
    return result



def default_config():
    user = Path(os.environ.get("USERPROFILE", str(Path.home())))
    one = os.environ.get("OneDrive")
    browser = Path(r"D:\浏览器2.lnk")
    if not browser.exists():
        browser = user / "Desktop" / "浏览器2.lnk"
    if one and not browser.exists():
        browser = Path(one) / "Desktop" / "浏览器2.lnk"
    return {
        "creator_url": "https://creator.douyin.com/",
        "browser_path": str(browser),
        "image_dir": r"D:\抖音\河北",
        "excel_path": r"D:\文案.xlsx",
        "sheet_name": "sheet1",
        "publish_date": "2026-06-21",
        "publish_start_date": "2026-06-21",
        "publish_end_date": "2026-06-21",
        "start_hour": 8,
        "end_hour": 18,
        "per_hour_count": 3,
        "custom_minutes": "5,25,45",
        "wait_min_seconds": 2,
        "wait_max_seconds": 5,
        "publish_interval_min_seconds": 30,
        "publish_interval_max_seconds": 90,
        "upload_check_interval_seconds": 3,
        "upload_max_wait_seconds": 180,
        "topic_wait_seconds": 3,
        "creator_center_wait_seconds": 8,
        "retry_times": 3,
        "overall_retry_times": 1,
        "cdp_port": 9222,
        "browser_user_data_dir": "",
        "browser_profile_directory": "",
        "browser_accounts": [],
        "browser_queue_column_widths": dict(BROWSER_QUEUE_DEFAULT_WIDTHS),
        "rotate_accounts_each_post": False,
        "rotate_batch_size": 1,
        "daily_limit_switch_wait_seconds": 10,
        "browser_account_suggestions": [
            {
                "name": "浏览器9",
                "browser_path": r"D:\浏览器9.lnk",
                "cdp_port": 9229,
                "posts_per_run": 1,
                "enabled": True,
                "close_after_finish": True,
            }
        ],
        "reuse_existing_cdp": True,
        "close_chrome_before_start": False,
        "no_raise_browser": True,
        "max_posts_this_run": 0,
        "state_path": r"D:\抖音\河北_publish_state.json",
        "log_path": r"D:\抖音\河北_publish_log.csv",
        "debug_dir": r"D:\抖音\debug_screenshots",
        "delete_image_after_success": True,
        "delete_copy_after_success": True,
        "use_schedule": True,
        "music_required": True,
        "keep_browser_open": True,
        "auto_switch_graphic_mode": True,
        "platform_music_only": True,
        "random_image": False,
        "weekly_start_enabled": False,
        "weekly_start_days": [0, 1, 2, 3, 4],
        "weekly_start_time": "09:00",
    }


def parse_bool(value, default=False):
    if isinstance(value, bool):
        return value
    if value is None:
        return bool(default)
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "on", "是", "启用"}:
        return True
    if text in {"0", "false", "no", "off", "否", "停用"}:
        return False
    return bool(default)


def normalize_browser_account(raw, index, legacy_cfg=None):
    """把旧单浏览器配置或新版账号配置规范化为同一种结构。"""
    raw = raw if isinstance(raw, dict) else {}
    legacy_cfg = legacy_cfg if isinstance(legacy_cfg, dict) else {}
    browser_path = str(raw.get("browser_path") or legacy_cfg.get("browser_path") or "").strip()
    name = str(raw.get("name") or Path(browser_path).stem or f"浏览器{index + 1}").strip()
    try:
        port = int(raw.get("cdp_port", legacy_cfg.get("cdp_port", 9222 + index)) or (9222 + index))
    except Exception:
        port = 9222 + index
    if not 1 <= port <= 65535:
        port = 9222 + index
    try:
        posts = int(raw.get("posts_per_run", 1) or 1)
    except Exception:
        posts = 1
    posts = max(1, posts)
    account_id = str(raw.get("id") or "").strip()
    if not account_id:
        digest = hashlib.sha1(f"{index}|{browser_path}|{name}".encode("utf-8", errors="ignore")).hexdigest()[:10]
        account_id = f"browser_{digest}"
    def inherited(key, fallback=""):
        return raw.get(key) if key in raw and raw.get(key) not in {None, ""} else legacy_cfg.get(key, fallback)

    try:
        start_hour = int(inherited("start_hour", 8))
        end_hour = int(inherited("end_hour", 18))
        per_hour_count = max(1, int(inherited("per_hour_count", 3)))
    except Exception:
        start_hour, end_hour, per_hour_count = 8, 18, 3
    return {
        "id": account_id,
        "name": name or f"浏览器{index + 1}",
        "enabled": parse_bool(raw.get("enabled"), True),
        "browser_path": browser_path,
        "browser_user_data_dir": str(
            raw.get("browser_user_data_dir", legacy_cfg.get("browser_user_data_dir", "")) or ""
        ).strip(),
        "browser_profile_directory": str(
            raw.get("browser_profile_directory", legacy_cfg.get("browser_profile_directory", "")) or ""
        ).strip(),
        "cdp_port": port,
        "posts_per_run": posts,
        "close_after_finish": parse_bool(raw.get("close_after_finish"), True),
        "independent_content": True,
        "excel_path": str(inherited("excel_path", "") or "").strip(),
        "sheet_name": str(inherited("sheet_name", "sheet1") or "sheet1").strip(),
        "use_schedule": parse_bool(inherited("use_schedule", True), True),
        "publish_start_date": str(
            inherited("publish_start_date", inherited("publish_date", "")) or ""
        ).strip(),
        "publish_end_date": str(
            inherited("publish_end_date", inherited("publish_start_date", inherited("publish_date", ""))) or ""
        ).strip(),
        "start_hour": start_hour,
        "end_hour": end_hour,
        "per_hour_count": per_hour_count,
        "custom_minutes": str(inherited("custom_minutes", "") or "").strip(),
    }


def new_browser_queue_id():
    """队列项必须有独立标识；同一浏览器可以在队列中出现多次。"""
    return f"browser_{uuid.uuid4().hex[:12]}"


def browser_runtime_identity(account):
    """判断两个队列项是否连接同一个浏览器运行实例。"""
    def normalized_path(value):
        text = str(value or "").strip()
        return os.path.normcase(os.path.normpath(text)) if text else ""

    return (
        normalized_path(account.get("browser_path")),
        normalized_path(account.get("browser_user_data_dir")),
        str(account.get("browser_profile_directory") or "").strip().casefold(),
    )


def validate_browser_queue_ports(accounts):
    """相同浏览器可复用端口；不同浏览器仍禁止端口冲突。"""
    port_owners = {}
    for account in accounts:
        port = int(account.get("cdp_port", 0) or 0)
        identity = browser_runtime_identity(account)
        previous = port_owners.get(port)
        if previous is not None and previous != identity:
            raise RuntimeError(
                f"CDP 调试端口 {port} 被不同浏览器队列项重复使用；"
                "同一浏览器可以重复排队，但不同浏览器必须使用不同端口。"
            )
        port_owners[port] = identity


def browser_accounts_from_config(cfg, enabled_only=False):
    raw_accounts = cfg.get("browser_accounts")
    accounts = []
    used_ids = set()
    if isinstance(raw_accounts, list):
        for index, raw in enumerate(raw_accounts):
            account = normalize_browser_account(raw, index, cfg)
            account_id = str(account.get("id") or "").strip()
            if account_id in used_ids:
                seed = f"{account_id}|{index}|{account.get('browser_path')}|{account.get('name')}"
                suffix = hashlib.sha1(seed.encode("utf-8", errors="ignore")).hexdigest()[:8]
                account_id = f"{account_id}_queue_{suffix}"
                account["id"] = account_id
            used_ids.add(account_id)
            accounts.append(account)
    if not accounts:
        accounts = [normalize_browser_account({}, 0, cfg)]
        existing_paths = {str(Path(accounts[0]["browser_path"])).lower()}
        suggestions = cfg.get("browser_account_suggestions")
        if isinstance(suggestions, list):
            for raw in suggestions:
                suggestion = normalize_browser_account(raw, len(accounts), cfg)
                suggestion_path = str(Path(suggestion["browser_path"])).lower()
                if suggestion_path in existing_paths or not Path(suggestion["browser_path"]).is_file():
                    continue
                accounts.append(suggestion)
                existing_paths.add(suggestion_path)
    if enabled_only:
        accounts = [account for account in accounts if account.get("enabled", True)]
    return accounts


def config_for_browser_account(cfg, account):
    result = copy.deepcopy(cfg)
    result["browser_path"] = account["browser_path"]
    result["browser_user_data_dir"] = account.get("browser_user_data_dir", "")
    result["browser_profile_directory"] = account.get("browser_profile_directory", "")
    result["cdp_port"] = int(account["cdp_port"])
    result["active_browser_account_id"] = account.get("id", "")
    result["active_browser_account_name"] = account.get("name", "")
    # 每个队列项都拥有独立进度。排期序号不会从前一个浏览器接着累计；
    # 同一个浏览器重复入队时，也按每一行队列任务分别计算。
    base_state_path = Path(
        str(cfg.get("_queue_state_base_path") or cfg.get("state_path") or "douyin_publish_state.json")
    )
    safe_id = re.sub(r"[^0-9A-Za-z_-]+", "_", str(account.get("id") or "account"))
    result["_queue_state_base_path"] = str(base_state_path)
    result["state_path"] = str(
        base_state_path.with_name(
            f"{base_state_path.stem}_{safe_id}{base_state_path.suffix or '.json'}"
        )
    )
    for key in (
        "excel_path", "sheet_name", "publish_start_date",
        "publish_end_date", "custom_minutes",
    ):
        result[key] = str(account.get(key) or "").strip()
    for key in ("start_hour", "end_hour", "per_hour_count"):
        result[key] = int(account.get(key, result.get(key, 0)) or 0)
    result["use_schedule"] = parse_bool(account.get("use_schedule"), True)
    return result


def load_config():
    cfg = default_config()
    source = CONFIG_PATH
    if IS_FROZEN and not source.exists():
        candidates = [
            INSTALL_DIR / "app" / "douyin_gui_config.json",
            INSTALL_DIR / "douyin_gui_config.json",
        ]
        source = next((path for path in candidates if path.is_file()), CONFIG_PATH)
    if source.exists():
        try:
            cfg.update(json.loads(source.read_text(encoding="utf-8")))
        except Exception:
            pass
    cfg["browser_accounts"] = browser_accounts_from_config(cfg)
    if IS_FROZEN and source != CONFIG_PATH:
        try:
            save_config(cfg)
        except Exception:
            pass
    return cfg


def save_config(cfg):
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


DEBUG_SCREENSHOT_NAME_RE = re.compile(r"^\d{8}_\d{6}_.+\.png$", re.IGNORECASE)


def maintain_debug_screenshots(cfg, state_path=None):
    """每启动 GUI 三次清理一次本软件生成的调试截图；不递归、不删除其他图片。"""
    state_path = Path(state_path or DEBUG_SCREENSHOT_CLEANUP_STATE_PATH)
    result = {"launch_count": 0, "cleaned": False, "removed": 0, "error": ""}
    try:
        try:
            state = json.loads(state_path.read_text(encoding="utf-8")) if state_path.is_file() else {}
        except Exception:
            state = {}
        launch_count = _nonnegative_state_int(state.get("launch_count"), 0) + 1
        if launch_count >= 3:
            debug_dir = Path(str(cfg.get("debug_dir", "")).strip().strip('"'))
            if debug_dir.is_dir():
                for path in debug_dir.iterdir():
                    if not path.is_file() or not DEBUG_SCREENSHOT_NAME_RE.match(path.name):
                        continue
                    try:
                        path.unlink()
                        result["removed"] += 1
                    except OSError:
                        pass
            launch_count = 0
            result["cleaned"] = True
        result["launch_count"] = launch_count
        state_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = state_path.with_name(state_path.name + f".{os.getpid()}.tmp")
        temp_path.write_text(
            json.dumps({
                "launch_count": launch_count,
                "last_launch_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(temp_path, state_path)
    except Exception as exc:
        result["error"] = repr(exc)
    return result


LEGACY_INSTALL_FILES = [
    "Start_Douyin_Publisher.bat",
    "start_app.vbs",
    "start_app.bat",
    "debug_start.bat",
    "requirements.txt",
    "version.json",
    "app/start_app.vbs",
    "app/start_app.bat",
    "app/debug_start.bat",
    "app/app_main.py",
    "app/app_main.pyw",
    "app/online_updater.py",
    "app/assets/app_logo.png",
    "app/assets/app_logo.ico",
    "packages/douyin_single_publish_v2.4.0.zip",
    "packages/douyin_single_publish_v2.3.0_no_black_window_update.zip",
    "packages/douyin_single_publish_v2.2.8_validate_copy_fix.zip",
    "packages/douyin_single_publish_v2.2.7_confirmed_update_ok.zip",
    "packages/douyin_single_publish_v2.2.6_keep_original_launcher_fix.zip",
    "packages/douyin_single_publish_v2.2.5_launcher_encoding_fix.zip",
    "packages/douyin_single_publish_v2.2.4_wps_py_fix.zip",
    "packages/douyin_single_publish_v2.2.3_encoding_fixed.zip",
    "packages/douyin_single_publish_v2.2.1_online_update_fixed.zip",
    "packages/douyin_single_publish_v2.2.0_online_update.zip",
    "packages/app_v2.1.zip",
]


def cleanup_legacy_install_after_migration():
    """EXE 首次启动并成功迁移配置后，清理明确的旧 Python/VBS 发行文件。"""
    if not IS_FROZEN or not CONFIG_PATH.is_file():
        return
    root = INSTALL_DIR.resolve()
    paths = list(LEGACY_INSTALL_FILES)
    legacy_config = root / "app" / "douyin_gui_config.json"
    if legacy_config.is_file():
        paths.append("app/douyin_gui_config.json")
    for rel in paths:
        target = (root / rel).resolve()
        try:
            target.relative_to(root)
        except ValueError:
            continue
        try:
            if target.is_file():
                target.unlink()
        except OSError:
            pass
    cache_dir = root / "app" / "__pycache__"
    if cache_dir.is_dir():
        shutil.rmtree(cache_dir, ignore_errors=True)
    for directory in (
        root / "app" / "assets",
        root / "app",
        root / "packages",
    ):
        try:
            directory.rmdir()
        except OSError:
            pass


def wlog(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def emit_browser_status(account, status, message="", success_count=0, quota=0):
    """向 GUI 回传当前浏览器账号状态，不混入普通日志内容。"""
    try:
        payload = {
            "id": str(account.get("id", "")),
            "name": str(account.get("name", "")),
            "status": str(status or ""),
            "message": str(message or ""),
            "success_count": int(success_count or 0),
            "quota": int(quota or 0),
        }
        print(BROWSER_STATUS_EVENT_PREFIX + json.dumps(payload, ensure_ascii=False), flush=True)
    except Exception:
        pass


def emit_countdown_event(action, seconds=0, reason="", kind="wait"):
    """向主界面发送倒计时状态；事件行不会写入普通日志框。"""
    try:
        payload = {
            "action": str(action),
            "seconds": max(0.0, float(seconds or 0)),
            "reason": str(reason or ""),
            "kind": str(kind or "wait"),
        }
        print(COUNTDOWN_EVENT_PREFIX + json.dumps(payload, ensure_ascii=False), flush=True)
    except Exception:
        pass


def countdown_sleep(cfg, seconds, reason="", kind="wait"):
    """可暂停、可回传剩余时间的等待，供 GUI 实时显示倒计时。"""
    try:
        remaining = max(0.0, float(seconds or 0))
    except Exception:
        remaining = 0.0
    if remaining <= 0:
        return

    total = remaining
    display_reason = str(reason or "等待下一步")
    emit_countdown_event("start", remaining, display_reason, kind)
    last_shown = int(math.ceil(remaining))
    try:
        while remaining > 0:
            if cfg:
                try:
                    if pause_flag_path(cfg).exists():
                        emit_countdown_event("pause", remaining, display_reason, kind)
                        wait_if_paused(cfg)
                        emit_countdown_event("resume", remaining, display_reason, kind)
                except Exception:
                    pass

            step = min(0.2, remaining)
            started = time.monotonic()
            time.sleep(step)
            elapsed = max(0.0, time.monotonic() - started)
            remaining = max(0.0, remaining - elapsed)
            shown = int(math.ceil(remaining))
            if shown != last_shown:
                last_shown = shown
                emit_countdown_event("tick", remaining, display_reason, kind)
    finally:
        emit_countdown_event("clear", 0, display_reason, kind)


def popup_pause_alert(msg="脚本已暂停，请查看前端日志。"):
    """
    脚本暂停/自动停止时弹出 Windows 提示框。
    只做提醒，不终止脚本进程。
    """
    try:
        ctypes.windll.user32.MessageBoxW(
            None,
            str(msg),
            APP_NAME,
            0x00000040 | 0x00001000
        )
    except Exception:
        pass


def pause_flag_path(cfg):
    try:
        state = Path(cfg.get("state_path", "douyin_publish_state.json"))
        return state.with_suffix(".pause")
    except Exception:
        return APP_DIR / "douyin_pause.flag"


def pause_flag_paths_for_queue(cfg, accounts=None):
    """GUI 与账号 worker 使用同一组暂停标记，停止或重启时可完整解除。"""
    paths = [pause_flag_path(cfg)]
    try:
        queue_accounts = accounts if accounts is not None else browser_accounts_from_config(cfg)
        paths.extend(pause_flag_path(config_for_browser_account(cfg, account)) for account in queue_accounts)
    except Exception:
        pass
    unique = []
    seen = set()
    for path in paths:
        key = os.path.normcase(os.path.abspath(str(path)))
        if key not in seen:
            seen.add(key)
            unique.append(Path(path))
    return unique


def clear_pause_flags_for_queue(cfg, accounts=None):
    for flag in pause_flag_paths_for_queue(cfg, accounts):
        try:
            if flag.exists():
                flag.unlink()
        except OSError:
            pass


def wait_if_paused(cfg):
    """
    前端点击"暂停/继续"时使用。
    注意：这里只等待暂停标记，不 terminate 进程，避免中途强断 Playwright 导致 EPIPE。
    """
    try:
        flag = pause_flag_path(cfg)
        shown = False
        while flag.exists():
            if not shown:
                msg = "脚本已暂停。点击前端'暂停/继续'后会继续运行。"
                wlog(msg)
                popup_pause_alert(msg)
                shown = True
            time.sleep(1)
        if shown:
            wlog("暂停已解除，继续运行。")
    except Exception:
        pass


def alert_auto_pause(msg):
    """所有自动暂停/保护性停止统一上报；无 GUI 父进程时直接弹窗。"""
    wlog(msg)
    print(AUTO_PAUSE_EVENT_PREFIX + json.dumps({"message": str(msg)}, ensure_ascii=False), flush=True)
    if os.environ.get("DOUYIN_GUI_PARENT") != "1":
        popup_pause_alert(msg)

def step_wait(cfg=None, reason=""):
    cfg = cfg or ACTIVE_CFG or {}
    if cfg:
        wait_if_paused(cfg)
    a = float(cfg.get("wait_min_seconds", 2) or 2)
    b = float(cfg.get("wait_max_seconds", 5) or 5)
    if a < 0:
        a = 0
    if b < a:
        b = a
    d = random.uniform(a, b) if b > a else a
    display_reason = reason or "步骤等待"
    if reason:
        wlog(f"等待 {d:.1f} 秒：{reason}")
    countdown_sleep(cfg, d, display_reason, "step_wait")


def worker_python_executable():
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        py = exe.with_name("python.exe")
        if py.exists():
            return str(py)
    return str(exe)


def application_command(*args):
    """源码模式使用系统 Python；冻结后子任务复用同一个 EXE。"""
    if IS_FROZEN:
        return [str(sys.executable), *[str(item) for item in args]]
    return [worker_python_executable(), str(Path(__file__).resolve()), *[str(item) for item in args]]


def no_window_creationflags():
    return getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0


def hidden_subprocess_kwargs():
    """Prevent console windows from flashing for helper processes on Windows."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = getattr(subprocess, "SW_HIDE", 0)
    return {
        "creationflags": no_window_creationflags(),
        "startupinfo": startupinfo,
    }


def hide_own_console_window():
    """Hide an accidentally attached console when the GUI is started with python.exe."""
    if os.name != "nt" or os.environ.get("DOUYIN_KEEP_CONSOLE") == "1":
        return
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)
    except Exception:
        pass


def run_ps(code):
    p = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", code],
                       capture_output=True, text=True, encoding="utf-8", errors="ignore",
                       **hidden_subprocess_kwargs())
    if p.returncode != 0:
        raise RuntimeError(p.stderr.strip() or p.stdout.strip())
    return p.stdout.strip()


def find_chrome_pids_for_cdp(port):
    """只返回命令行明确包含目标 CDP 端口的 Chrome 主进程。"""
    port = int(port)
    ps = rf"""
    $items = @(Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
      Where-Object {{ $_.CommandLine -match '(?i)--remote-debugging-port(?:=|\s+)"?{port}(?:\D|$)' }} |
      Select-Object -ExpandProperty ProcessId)
    $items | ConvertTo-Json -Compress
    """
    try:
        output = run_ps(ps).strip()
        if not output:
            return []
        data = json.loads(output)
        values = data if isinstance(data, list) else [data]
        return sorted({int(value) for value in values if int(value) > 0})
    except Exception:
        return []


def close_browser_launched_by_app(browser, launch_info, display_name="浏览器"):
    """先走 CDP 正常关闭；仍存活时仅结束本次端口对应的精确进程树。"""
    if not launch_info or not launch_info.get("launched"):
        return False
    port = int(launch_info.get("port", 0) or 0)
    try:
        if browser is not None:
            browser.close()
    except Exception:
        pass
    if not port or not wait_cdp(port, 1.5):
        wlog(f"已关闭本软件启动的 {display_name}。")
        return True
    pids = list(launch_info.get("chrome_pids") or []) or find_chrome_pids_for_cdp(port)
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/PID", str(int(pid)), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                **hidden_subprocess_kwargs(),
            )
        except Exception:
            pass
    if wait_cdp(port, 2):
        wlog(f"提醒：{display_name} 的端口 {port} 仍在监听，未按名称强制关闭其它 Chrome。")
        return False
    wlog(f"已关闭本软件启动的 {display_name}（端口 {port}）。")
    return True


def resolve_lnk(path):
    ps = rf"""
    $s=(New-Object -COM WScript.Shell).CreateShortcut('{str(path)}');
    [PSCustomObject]@{{TargetPath=$s.TargetPath; Arguments=$s.Arguments}} | ConvertTo-Json -Compress
    """
    data = json.loads(run_ps(ps))
    return data.get("TargetPath", ""), data.get("Arguments", "")


def wait_cdp(port, seconds=25):
    url = f"http://127.0.0.1:{port}/json/version"
    end = time.time() + seconds
    while time.time() < end:
        try:
            with urllib.request.urlopen(url, timeout=1.2) as r:
                if r.status == 200:
                    return True
        except Exception:
            time.sleep(0.6)
    return False


def safe_bring_to_front(cfg, page):
    """
    后台运行模式：默认不主动拉起/置顶浏览器。
    需要看窗口时，把配置 no_raise_browser 改为 False。
    """
    try:
        if cfg and cfg.get("no_raise_browser", True):
            return
        page.bring_to_front()
    except Exception:
        pass


def sanitize_chrome_args(args):
    """删除快捷方式里可能冲突的调试端口参数。"""
    args = str(args or "")
    patterns = [
        r"\s*--remote-debugging-port(?:=|\s+)\"[^\"]*\"",
        r"\s*--remote-debugging-port(?:=|\s+)\S+",
        r"\s*--remote-debugging-address(?:=|\s+)\"[^\"]*\"",
        r"\s*--remote-debugging-address(?:=|\s+)\S+",
    ]
    for pat in patterns:
        args = re.sub(pat, "", args, flags=re.I)
    return args.strip()


def strip_profile_args(args):
    """当用户在前端指定用户目录/Profile时，删除快捷方式里旧的同类参数，避免重复。"""
    args = str(args or "")
    patterns = [
        r"\s*--user-data-dir(?:=|\s+)\"[^\"]*\"",
        r"\s*--user-data-dir(?:=|\s+)\S+",
        r"\s*--profile-directory(?:=|\s+)\"[^\"]*\"",
        r"\s*--profile-directory(?:=|\s+)\S+",
    ]
    for pat in patterns:
        args = re.sub(pat, "", args, flags=re.I)
    return args.strip()


def has_user_data_dir_arg(args):
    return "--user-data-dir" in str(args or "").lower()


def kill_chrome_residue(cfg):
    if not cfg.get("close_chrome_before_start", False):
        return
    wlog("已启用：启动前关闭 Chrome / node 残留进程。")
    for proc in ["chrome.exe", "node.exe"]:
        try:
            subprocess.run(["taskkill", "/F", "/IM", proc], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           **hidden_subprocess_kwargs())
        except Exception:
            pass

def build_browser_launch_command(target, args, port, target_url, headless=False):
    """构造浏览器启动命令；轮换子任务可使用真正无窗口的 Chromium Headless 模式。"""
    args = str(args or "").strip()
    if headless:
        # Chromium/Chrome/Edge 的新版无头模式：不会创建桌面浏览器窗口，
        # 但仍保留完整页面渲染、上传文件、CDP 控制和持久化用户目录能力。
        args = (args + ' --headless=new --window-size=1440,1100 --disable-gpu').strip()
    return f'"{target}" {args} --remote-debugging-port={int(port)} "{target_url}"'.strip()


def open_browser_with_cdp(cfg, force_new=False):
    """
    v2.1：稳定 CDP 后台版。
    - 用户不用改现有浏览器快捷方式目标；脚本会在启动时临时追加 --remote-debugging-port。
    - 可在前端填写 browser_user_data_dir / browser_profile_directory，脚本会临时追加，不修改 .lnk 文件本身。
    - 自动清理重复的 --remote-debugging-port，避免快捷方式和脚本重复写端口。
    - 当运行时配置 _headless_browser=True 时，使用 Chromium 新版 Headless，桌面不弹出浏览器窗口。
    """
    port = int(cfg.get("cdp_port", 9222))
    kill_chrome_residue(cfg)

    if wait_cdp(port, 2):
        if cfg.get("reuse_existing_cdp", True):
            wlog(f"已检测到浏览器调试端口 {port}，复用现有调试连接。")
            return {"launched": False, "process": None, "port": port, "reused": True}
        raise RuntimeError(f"调试端口 {port} 已被占用。请关闭旧 Chrome，或勾选'启动前关闭Chrome残留'。")

    browser_path = Path(cfg["browser_path"])
    if not browser_path.exists():
        raise FileNotFoundError(f"浏览器路径不存在：{browser_path}")

    target_url = cfg.get("creator_url", "https://creator.douyin.com/")
    target = str(browser_path)
    args = ""

    if browser_path.suffix.lower() == ".lnk":
        target, args = resolve_lnk(browser_path)
        if not target:
            raise RuntimeError(f"无法解析浏览器快捷方式：{browser_path}")

    args = sanitize_chrome_args(args)

    user_dir = str(cfg.get("browser_user_data_dir", "") or "").strip()
    profile_dir = str(cfg.get("browser_profile_directory", "") or "").strip()

    # 如果前端指定了用户目录/Profile，就以这个为准；不会修改原快捷方式，只是本次启动临时追加参数。
    if user_dir:
        args = strip_profile_args(args)
        args = (args + f' --user-data-dir="{user_dir}"').strip()
        if profile_dir:
            args = (args + f' --profile-directory="{profile_dir}"').strip()
    else:
        # 如果快捷方式本身没有 user-data-dir，为了适配新版 Chrome 远程调试限制，使用脚本专用默认目录。
        if not has_user_data_dir_arg(args):
            default_profile = str(Path(os.environ.get("USERPROFILE", "D:")) / "douyin_chrome_profile_v21")
            args = (args + f' --user-data-dir="{default_profile}"').strip()
            wlog(f"快捷方式未带 user-data-dir，已临时使用脚本专用目录：{default_profile}")

    headless = parse_bool(cfg.get("_headless_browser"), False)
    cmd = build_browser_launch_command(target, args, port, target_url, headless=headless)
    if headless:
        wlog("轮换后台模式：浏览器以 Headless 方式启动，不会在桌面弹出窗口。")
    else:
        wlog("启动浏览器命令已生成：不会修改你的浏览器快捷方式目标，只临时追加参数。")
        wlog(f"打开浏览器：{target}")
    proc = subprocess.Popen(cmd, shell=True, **hidden_subprocess_kwargs())
    if not wait_cdp(port, 45):
        raise RuntimeError("浏览器已尝试启动，但无法连接调试端口。请检查用户目录是否被其它 Chrome 占用，或勾选启动前关闭Chrome残留。")
    return {
        "launched": True,
        "process": proc,
        "port": port,
        "reused": False,
        "chrome_pids": find_chrome_pids_for_cdp(port),
    }



def read_state(cfg):
    p = Path(cfg["state_path"])
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"copy_index": 0, "slot_index": 0}


def write_state(cfg, state):
    p = Path(cfg["state_path"])
    p.parent.mkdir(parents=True, exist_ok=True)
    temp_path = p.with_name(f".{p.name}.{os.getpid()}.{threading.get_ident()}.tmp")
    try:
        temp_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temp_path, p)
    finally:
        try:
            if temp_path.exists():
                temp_path.unlink()
        except OSError:
            pass


def _nonnegative_state_int(value, default=0):
    try:
        return max(0, int(value or 0))
    except Exception:
        return max(0, int(default or 0))


def account_run_resume_progress(state, queue_run_id, quota):
    """恢复同一队列轮次的账号进度，并兼容 2.5.1 仅有 slot_index 的状态。"""
    state = state if isinstance(state, dict) else {}
    quota = max(1, _nonnegative_state_int(quota, 1))
    stored_run_id = str(state.get("queue_run_id") or "").strip()
    current_run_id = str(queue_run_id or "").strip()
    if stored_run_id and current_run_id and stored_run_id == current_run_id:
        return min(quota, _nonnegative_state_int(state.get("run_progress"), 0))
    if stored_run_id or "run_progress" in state:
        return 0
    # 旧版没有本轮进度字段。slot_index 每次成功后才递增，因此可用于恢复未完成配额。
    return _nonnegative_state_int(state.get("slot_index"), 0) % quota


def queue_run_signature(accounts):
    return "|".join(
        f"{str(account.get('id') or '')}:{max(1, _nonnegative_state_int(account.get('posts_per_run'), 1))}"
        for account in accounts
    )


def prepare_queue_run(cfg, accounts):
    """开始新队列轮次或恢复上次被停止的轮次。"""
    state = read_state(cfg)
    account_ids = {str(account.get("id") or "") for account in accounts}
    active = parse_bool(state.get("queue_run_active"), False)
    run_id = str(state.get("queue_run_id") or "").strip()
    if not active or not run_id:
        run_id = f"queue_{uuid.uuid4().hex}"
        completed = set()
        resumed = False
    else:
        completed = {
            str(account_id) for account_id in (state.get("queue_completed_account_ids") or [])
            if str(account_id) in account_ids
        }
        resumed = True
    state.update({
        "queue_run_id": run_id,
        "queue_run_active": True,
        "queue_run_signature": queue_run_signature(accounts),
        "queue_completed_account_ids": sorted(completed),
    })
    write_state(cfg, state)
    return run_id, completed, resumed


def update_queue_run_state(cfg, run_id, accounts, completed_ids, active):
    state = read_state(cfg)
    state.update({
        "queue_run_id": str(run_id or ""),
        "queue_run_active": bool(active),
        "queue_run_signature": queue_run_signature(accounts),
        "queue_completed_account_ids": sorted({str(item) for item in completed_ids if str(item)}),
    })
    if not active:
        state["queue_run_completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    write_state(cfg, state)


def append_log(cfg, row):
    p = Path(cfg["log_path"])
    p.parent.mkdir(parents=True, exist_ok=True)
    exists = p.exists()
    with p.open("a", newline="", encoding="utf-8-sig") as f:
        fields = ["time", "image", "copy_index", "schedule_time", "copy_preview", "status"]
        wr = csv.DictWriter(f, fieldnames=fields)
        if not exists:
            wr.writeheader()
        wr.writerow(row)



def schedule_signature(cfg):
    """
    v2.1：排期签名支持多日定时。
    只要日期范围、小时、分钟点、每小时数量或定时开关变化，就重置 slot_index。
    """
    keys = [
        "publish_start_date",
        "publish_end_date",
        "publish_date",
        "start_hour",
        "end_hour",
        "custom_minutes",
        "per_hour_count",
        "use_schedule",
    ]
    return "|".join(f"{k}={str(cfg.get(k, ''))}" for k in keys)

def build_slots(cfg, log_result=True):
    """
    v2.1：支持一次运行多日定时。
    新字段：publish_start_date / publish_end_date。
    兼容旧字段 publish_date：如果没有填写开始/结束日期，就按旧字段单日排期。
    """
    from datetime import timedelta

    start_date_raw = str(cfg.get("publish_start_date") or cfg.get("publish_date") or "").strip()
    end_date_raw = str(cfg.get("publish_end_date") or start_date_raw).strip()
    if not start_date_raw:
        raise RuntimeError("开始日期不能为空，例如 2026-06-21。")
    if not end_date_raw:
        end_date_raw = start_date_raw

    try:
        start_date = datetime.strptime(start_date_raw, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_date_raw, "%Y-%m-%d").date()
    except Exception:
        raise RuntimeError("开始日期/结束日期格式必须是 YYYY-MM-DD，例如 2026-06-21。")

    if end_date < start_date:
        raise RuntimeError("结束日期不能早于开始日期。")

    try:
        start = int(cfg.get("start_hour", 8))
        end = int(cfg.get("end_hour", 18))
    except Exception:
        raise RuntimeError("开始小时和结束小时必须填写数字，例如 8、18、23、24。")

    if start < 0 or start > 23:
        raise RuntimeError("开始小时只能填写 0-23，例如 8 或 17。")
    if end < 0 or end > 24:
        raise RuntimeError("结束小时只能填写 0-24。填 24 表示当天 24 点前。")

    end_for_range = 23 if end == 24 else end
    per = max(1, int(cfg.get("per_hour_count", 3)))
    raw = str(cfg.get("custom_minutes", "")).strip()
    if raw:
        mins = []
        for x in re.split(r"[，,\s]+", raw):
            if not x.strip():
                continue
            mi = int(x)
            if mi < 0 or mi > 59:
                raise RuntimeError(f"分钟点只能填写 0-59，当前填写了：{mi}")
            mins.append(mi)
        mins = sorted(set(mins))
    else:
        step = 60 // per
        mins = [min(59, 5 + i * step) for i in range(per)]

    if not mins:
        raise RuntimeError("分钟点不能为空，例如填写：5,25,45")

    slots = []
    day = start_date
    while day <= end_date:
        if end_for_range >= start:
            for h in range(start, end_for_range + 1):
                for mi in mins[:per]:
                    slots.append(datetime(day.year, day.month, day.day, h, mi))
        else:
            # 跨夜排期：例如 20 到 2，会生成当天 20-23 + 次日 0-2。
            for h in range(start, 24):
                for mi in mins[:per]:
                    slots.append(datetime(day.year, day.month, day.day, h, mi))
            next_day = day + timedelta(days=1)
            for h in range(0, end_for_range + 1):
                for mi in mins[:per]:
                    slots.append(datetime(next_day.year, next_day.month, next_day.day, h, mi))
        day = day + timedelta(days=1)

    if not slots:
        raise RuntimeError("没有生成任何定时时间，请检查日期、开始小时、结束小时、每小时条数和分钟点。")
    if log_result:
        wlog(f"已生成多日排期 {len(slots)} 个：从 {slots[0].strftime('%Y-%m-%d %H:%M')} 到 {slots[-1].strftime('%Y-%m-%d %H:%M')}")
    return slots


def repeating_schedule_slot(slots, progress_index):
    """时间点用完后从本轮第一个时间点继续，返回时间点、轮次和轮内序号。"""
    if not slots:
        raise RuntimeError("定时发布时间点为空，无法循环使用。")
    try:
        progress = max(0, int(progress_index or 0))
    except Exception:
        progress = 0
    cycle_index, slot_offset = divmod(progress, len(slots))
    return slots[slot_offset], cycle_index + 1, slot_offset


IMAGE_FILE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}


def list_images(cfg, image_dir=None):
    d = Path(image_dir if image_dir is not None else cfg["image_dir"])
    if not d.exists():
        raise FileNotFoundError(f"图片文件夹不存在：{d}")
    if not d.is_dir():
        raise NotADirectoryError(f"图片路径不是文件夹：{d}")
    arr = [p for p in d.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_FILE_EXTS]
    arr.sort(key=lambda p: p.name.casefold())
    return arr


COPY_FILE_EXTS = {".xlsx", ".xls", ".xlsm"}
COPY_HEADER_NAMES = {
    "文案", "正文", "内容", "发布文案", "抖音文案", "作品文案",
    "文案内容", "正文内容", "发布内容", "内容文案", "作品内容",
    "发布正文", "文案正文", "文案栏", "正文栏", "内容栏",
    "文案列", "正文列", "内容列",
}
TITLE_HEADER_NAMES = {
    "标题", "作品标题", "发布标题", "抖音标题", "标题栏", "标题列", "标题内容",
}
CITY_HEADER_NAMES = {
    "城市", "地区", "所在城市", "发布城市", "城市栏", "城市名称",
    "城市名", "城市列", "素材城市", "图片城市", "地市", "所属城市",
    "城市地区", "城市区域", "省市",
}

BROWSER_QUEUE_TABLE_COLUMNS = [
    ("顺序", "_order"), ("队列ID", "id"), ("账号名称", "name"),
    ("启用", "enabled"), ("浏览器路径", "browser_path"),
    ("用户数据目录", "browser_user_data_dir"), ("Profile目录", "browser_profile_directory"),
    ("CDP端口", "cdp_port"), ("发布条数", "posts_per_run"),
    ("完成后关闭浏览器", "close_after_finish"),
    ("文案Excel", "excel_path"), ("工作表名", "sheet_name"),
    ("使用定时发布", "use_schedule"), ("开始日期", "publish_start_date"),
    ("结束日期", "publish_end_date"), ("开始小时", "start_hour"),
    ("结束小时", "end_hour"), ("每小时条数", "per_hour_count"),
    ("分钟点", "custom_minutes"),
]


def validate_copy_file_path(p):
    """防止把 .py 等脚本文件误当作文案表格读取。"""
    p = Path(str(p or "").strip().strip('"'))
    if not p:
        raise RuntimeError("文案文件未设置：请在前端选择 .xlsx / .xls / .xlsm 文案文件。")
    if not p.exists():
        raise FileNotFoundError(f"文案文件不存在：{p}")
    ext = p.suffix.lower()
    if ext not in COPY_FILE_EXTS:
        raise RuntimeError(
            "文案文件格式不支持。\n"
            f"当前选择：{p}\n"
            f"当前后缀：{ext or '无后缀'}\n"
            "请重新选择 .xlsx / .xls / .xlsm 文案文件。不要选择 .py、.txt 或 .csv 文件。"
        )
    return p


def read_plain_copy_file(p):
    # txt/csv 兜底读取，每行一条；csv 如果有“文案”列，优先读取文案列。
    if p.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "utf-8", "gbk"):
            try:
                df = pd.read_csv(p, encoding=enc).dropna(how="all")
                if not df.empty:
                    col = next((c for c in df.columns if "文案" in str(c)), df.columns[0])
                    return [str(v).strip() for v in df[col].tolist() if not pd.isna(v) and str(v).strip()]
            except Exception:
                pass
    for enc in ("utf-8-sig", "utf-8", "gbk"):
        try:
            return [line.strip() for line in p.read_text(encoding=enc).splitlines() if line.strip()]
        except Exception:
            continue
    raise RuntimeError(f"无法读取文案文本文件：{p}")


def copy_cell_text(value):
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    return str(value).strip()


def normalized_header_name(value):
    return re.sub(r"[\s_\-/（）()：:]+", "", copy_cell_text(value)).casefold()


def is_copy_header(value):
    return normalized_header_name(value) in COPY_HEADER_NAMES


def is_title_header(value):
    return normalized_header_name(value) in TITLE_HEADER_NAMES


def is_city_header(value):
    return normalized_header_name(value) in CITY_HEADER_NAMES


def normalized_city_name(value):
    """用于匹配 Excel 城市和图片子文件夹；兼容“济南”与“济南市”。"""
    name = re.sub(r"\s+", "", copy_cell_text(value)).casefold()
    if len(name) > 1 and name.endswith("市"):
        name = name[:-1]
    return name


def image_city_folder_names(cfg):
    """只读取图片总目录的直接子文件夹，不递归扫描用户的其他文件。"""
    root = Path(str(cfg.get("image_dir", "")).strip().strip('"'))
    if not root.is_dir():
        return []
    try:
        return [child.name for child in root.iterdir() if child.is_dir()]
    except OSError:
        return []


def infer_city_column(frame, known_city_names, excluded_columns=()):
    """依据图片子文件夹名称识别无表头/非常用表头中的城市列。"""
    known_keys = {normalized_city_name(name) for name in (known_city_names or [])}
    known_keys.discard("")
    if not known_keys or frame is None or frame.empty:
        return None

    excluded = set(excluded_columns)
    candidates = []
    for column in frame.columns:
        if column in excluded:
            continue
        values = [copy_cell_text(value) for value in frame[column].tolist()]
        values = [value for value in values if value]
        if not values:
            continue
        matched = sum(normalized_city_name(value) in known_keys for value in values)
        # 至少六成非空值应当是现有城市文件夹；错误项随后会给出具体城市提示。
        if matched and matched * 5 >= len(values) * 3:
            candidates.append((matched / len(values), matched, column))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidates[0][2]


def extract_publish_records(df, known_city_names=None):
    """
    读取发布记录。

    新格式使用“城市 / 标题 / 文案”表头；标题可以为空。存在城市列时，
    每条记录必须填写城市，并从图片总目录下的同名子文件夹选择图片。
    旧版无表头单列继续把第一行当作第一条文案；无表头多列可按
    “城市 / 文案”或“城市 / 标题 / 文案”读取；旧版仅以“标题”作为
    文案列表头的单列表格也继续兼容。
    """
    if df is None:
        return []
    frame = df.dropna(how="all")
    if frame.empty:
        return []

    first_index = frame.index[0]
    header_values = {column: copy_cell_text(frame.at[first_index, column]) for column in frame.columns}
    copy_column = next((column for column, value in header_values.items() if is_copy_header(value)), None)
    title_column = next((column for column, value in header_values.items() if is_title_header(value)), None)
    city_column = next((column for column, value in header_values.items() if is_city_header(value)), None)
    nonempty_header_columns = [column for column, value in header_values.items() if value]

    structured = copy_column is not None
    legacy_title_header = (
        not structured
        and title_column is not None
        and city_column is None
        and len(nonempty_header_columns) == 1
    )

    if structured:
        data_frame = frame.iloc[1:]
        if city_column is None:
            city_column = infer_city_column(
                data_frame,
                known_city_names,
                excluded_columns=(copy_column, title_column),
            )
        data_format = "structured"
    elif legacy_title_header:
        # 兼容历史上把单列文案表头命名为“标题”的文件。
        copy_column = title_column
        title_column = None
        data_frame = frame.iloc[1:]
        data_format = "legacy_header"
    elif title_column is not None or city_column is not None:
        raise RuntimeError("Excel 已识别到“城市/标题”表头，但缺少“文案”列。请增加文案表头。")
    else:
        active_columns = [
            column for column in frame.columns
            if any(copy_cell_text(value) for value in frame[column].tolist())
        ]
        city_column = None
        if len(active_columns) >= 2:
            city_column = infer_city_column(frame, known_city_names)
        if city_column is None:
            # 未识别到城市时保持旧版行为：始终读取第一列，不能因旁边有备注列而改变文案。
            copy_column = active_columns[0] if active_columns else frame.columns[0]
            title_column = None
        else:
            content_columns = [column for column in active_columns if column != city_column]
            copy_column = content_columns[-1] if content_columns else frame.columns[0]
            title_column = content_columns[-2] if len(content_columns) >= 2 else None
        data_frame = frame
        data_format = "headerless_city" if city_column is not None else "legacy_headerless"

    records = []
    for row_index, row in data_frame.iterrows():
        copy_text = copy_cell_text(row.get(copy_column))
        if not copy_text:
            continue
        try:
            excel_row = int(row_index) + 1
        except Exception:
            excel_row = 0
        records.append({
            "city": copy_cell_text(row.get(city_column)) if city_column is not None else "",
            "title": copy_cell_text(row.get(title_column)) if title_column is not None else "",
            "copy": copy_text,
            "excel_row": excel_row,
            "has_city_column": city_column is not None,
            "format": data_format,
        })
    return records


def extract_copy_texts(df):
    """同时兼容有表头和无表头；无表头时第一行就是第一条文案。"""
    return [record["copy"] for record in extract_publish_records(df)]


def read_publish_records(cfg):
    p = validate_copy_file_path(cfg.get("excel_path", ""))
    suffix = p.suffix.lower()
    # v2.2.7：不要让 pandas 自己猜 Excel 类型，按文件后缀指定引擎。
    engine = "openpyxl" if suffix in {".xlsx", ".xlsm"} else "xlrd"

    try:
        xls = pd.ExcelFile(p, engine=engine)
    except Exception as e:
        raise RuntimeError(
            "无法读取文案Excel。\n"
            f"当前选择：{p}\n"
            "请确认这是标准 Excel 文件，不是 .py/.txt 改后缀，也不要被 WPS/Excel 占用。\n"
            "如仍失败，请用 Excel/WPS 打开后另存为 Excel 工作簿（*.xlsx）。\n"
            f"原始错误：{repr(e)}"
        ) from e

    try:
        sheet_names = list(xls.sheet_names)
        sheet = next((s for s in sheet_names if s.lower() == str(cfg["sheet_name"]).lower()), None)
        if not sheet:
            raise RuntimeError(f"找不到工作表：{cfg['sheet_name']}；当前工作表：{sheet_names}")
        # 必须用 header=None 原样读取。否则无表头表格的第一条正文会被 pandas 当成列名。
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
    except Exception as e:
        if isinstance(e, RuntimeError) and str(e).startswith("找不到工作表："):
            raise
        raise RuntimeError(
            "读取工作表失败。\n"
            f"当前选择：{p}\n"
            f"当前工作表：{sheet}\n"
            "请尝试将表格另存为新的 .xlsx 文件，或新建空白表格复制文案列后再保存。\n"
            f"原始错误：{repr(e)}"
        ) from e
    finally:
        xls.close()

    return extract_publish_records(df, known_city_names=image_city_folder_names(cfg))


def read_copies(cfg):
    """兼容旧调用：只返回正文列表。新发布流程使用 read_publish_records。"""
    return [record["copy"] for record in read_publish_records(cfg)]


def normalize_copy_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def copy_key(value):
    """
    V28：文案唯一标记。用于 Excel 被占用无法删除时，临时跳过已发布文案，避免重复发。
    """
    text = normalize_copy_text(value)
    return hashlib.sha1(text.encode("utf-8", errors="ignore")).hexdigest()


def publish_record_key(record):
    payload = {
        "city": normalize_copy_text(record.get("city", "")),
        "title": normalize_copy_text(record.get("title", "")),
        "copy": normalize_copy_text(record.get("copy", "")),
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()


def filter_used_publish_records(records, state):
    used = set(state.get("used_copy_keys", []) or [])
    if not used:
        return records
    filtered = [
        record for record in records
        if publish_record_key(record) not in used and copy_key(record.get("copy", "")) not in used
    ]
    skipped = len(records) - len(filtered)
    if skipped:
        wlog(f"已根据进度文件跳过 {skipped} 条已发布但未能从 Excel 删除的记录。")
    return filtered


def filter_used_copies(copies, state):
    used = set(state.get("used_copy_keys", []) or [])
    if not used:
        return copies
    filtered = [c for c in copies if copy_key(c) not in used]
    skipped = len(copies) - len(filtered)
    if skipped:
        wlog(f"已根据进度文件跳过 {skipped} 条已发布但未能从 Excel 删除的文案。")
    return filtered


def mark_publish_record_used_in_state(cfg, record):
    state = read_state(cfg)
    used = list(state.get("used_copy_keys", []) or [])
    key = publish_record_key(record)
    if key not in used:
        used.append(key)
    state["used_copy_keys"] = used[-5000:]
    write_state(cfg, state)
    wlog("已把当前城市、标题和文案写入进度文件，避免后续重复使用。")


def mark_copy_used_in_state(cfg, copy_text):
    state = read_state(cfg)
    used = list(state.get("used_copy_keys", []) or [])
    key = copy_key(copy_text)
    if key not in used:
        used.append(key)
    state["used_copy_keys"] = used[-5000:]
    write_state(cfg, state)
    wlog("已把当前文案写入进度文件的已用文案列表，避免后续重复使用。")


def validate_city_folder_name(city):
    city = copy_cell_text(city)
    if not city:
        raise RuntimeError("当前文案的城市为空；使用城市表头时，每条文案都必须填写城市。")
    if city in {".", ".."} or any(ch in city for ch in ("/", "\\", ":", "\x00")):
        raise RuntimeError(f"城市名称不能包含路径字符：{city}")
    return city


def resolve_image_dir_for_record(cfg, record):
    root = Path(str(cfg.get("image_dir", "")).strip().strip('"'))
    if not root.is_dir():
        raise FileNotFoundError(f"图片总文件夹不存在：{root}")
    if not record.get("has_city_column", False):
        return root

    city = validate_city_folder_name(record.get("city", ""))
    city_key = normalized_city_name(city)
    matches = [
        child for child in root.iterdir()
        if child.is_dir() and normalized_city_name(child.name) == city_key
    ]
    if not matches:
        raise FileNotFoundError(f"找不到城市“{city}”对应的图片子文件夹：{root / city}")
    exact = next((child for child in matches if child.name == city), matches[0])
    root_resolved = root.resolve()
    exact_resolved = exact.resolve()
    if exact_resolved.parent != root_resolved:
        raise RuntimeError(f"城市图片目录必须位于图片总文件夹内：{exact}")
    return exact


def list_images_for_record(cfg, record):
    image_dir = resolve_image_dir_for_record(cfg, record)
    images = list_images(cfg, image_dir=image_dir)
    if not images:
        city = copy_cell_text(record.get("city", ""))
        scope = f"城市“{city}”" if record.get("has_city_column", False) else "图片根目录"
        if not record.get("has_city_column", False) and image_city_folder_names(cfg):
            raise RuntimeError(
                "图片总目录中已找到城市子文件夹，但 Excel 未识别到城市列。\n"
                "请使用表头“城市｜标题｜文案”（标题可不填），或无表头按相同列顺序填写。\n"
                f"图片总目录：{image_dir}"
            )
        raise RuntimeError(f"{scope}没有可发布图片：{image_dir}")
    return images


def choose_image_for_record(cfg, record):
    images = list_images_for_record(cfg, record)
    return random.choice(images) if cfg.get("random_image", False) else images[0]


def validate_publish_records_and_images(cfg):
    records = read_publish_records(cfg)
    if not records:
        raise RuntimeError("Excel 中没有可发布文案。")
    checked_scopes = set()
    for record in records:
        scope = (
            bool(record.get("has_city_column", False)),
            copy_cell_text(record.get("city", "")).casefold(),
        )
        if scope in checked_scopes:
            continue
        list_images_for_record(cfg, record)
        checked_scopes.add(scope)
    return records


def delete_copy_from_excel(cfg, copy_record):
    """
    V28：发布成功后从 Excel 中删除已使用文案所在行。
    如果 Excel 被 WPS/Excel 打开占用，会明确提示，并返回 False。
    """
    p = Path(cfg["excel_path"])
    if not p.exists():
        wlog(f"提醒：Excel 不存在，无法删除已用文案：{p}")
        return False

    record = copy_record if isinstance(copy_record, dict) else {
        "city": "", "title": "", "copy": copy_record, "excel_row": 0,
        "has_city_column": False, "format": "legacy",
    }
    target = normalize_copy_text(record.get("copy", ""))
    if not target:
        wlog("提醒：当前文案为空，跳过删除 Excel 文案。")
        return False

    try:
        if p.suffix.lower() == ".xls":
            wlog("旧版 .xls 不支持安全原地删行，已改用进度文件记录已用文案；建议另存为 .xlsx。")
            return False
        wb = load_workbook(p, keep_vba=(p.suffix.lower() == ".xlsm"))
    except PermissionError:
        wlog(f"删除文案失败：Excel 文件被占用，无法打开写入：{p}。请关闭 WPS/Excel 里的这个文件。")
        return False
    except Exception as e:
        wlog(f"删除文案失败：无法打开 Excel：{repr(e)}")
        return False

    sheet_name = next((s for s in wb.sheetnames if s.lower() == str(cfg["sheet_name"]).lower()), None)
    if not sheet_name:
        wlog(f"提醒：找不到工作表，无法删除已用文案：{cfg.get('sheet_name')}")
        wb.close()
        return False

    ws = wb[sheet_name]

    copy_col = next((c for c in range(1, ws.max_column + 1) if is_copy_header(ws.cell(row=1, column=c).value)), None)
    title_col = next((c for c in range(1, ws.max_column + 1) if is_title_header(ws.cell(row=1, column=c).value)), None)
    city_col = next((c for c in range(1, ws.max_column + 1) if is_city_header(ws.cell(row=1, column=c).value)), None)

    nonempty_header_cols = [
        c for c in range(1, ws.max_column + 1)
        if normalize_copy_text(ws.cell(row=1, column=c).value)
    ]
    legacy_title_header = copy_col is None and title_col is not None and city_col is None and len(nonempty_header_cols) == 1
    if legacy_title_header:
        copy_col = title_col
        title_col = None

    if copy_col is None:
        for c in range(1, ws.max_column + 1):
            if any(normalize_copy_text(ws.cell(row=r, column=c).value) for r in range(1, min(ws.max_row, 20) + 1)):
                copy_col = c
                break

    if copy_col is None:
        wlog("提醒：Excel 中没有找到可删除的文案列。")
        wb.close()
        return False

    has_header = is_copy_header(ws.cell(row=1, column=copy_col).value) or legacy_title_header
    candidate_rows = list(range(2 if has_header else 1, ws.max_row + 1))

    target_city = normalize_copy_text(record.get("city", ""))
    target_title = normalize_copy_text(record.get("title", ""))

    def row_matches(row_number):
        value = normalize_copy_text(ws.cell(row=row_number, column=copy_col).value)
        if value != target:
            return False
        if city_col is not None and normalize_copy_text(ws.cell(row=row_number, column=city_col).value) != target_city:
            return False
        if title_col is not None and normalize_copy_text(ws.cell(row=row_number, column=title_col).value) != target_title:
            return False
        return True

    matched_row = None
    preferred_row = int(record.get("excel_row", 0) or 0)
    if preferred_row in candidate_rows and row_matches(preferred_row):
        matched_row = preferred_row
    for r in candidate_rows:
        if matched_row is None and row_matches(r):
            matched_row = r
            break

    if matched_row is None and not isinstance(copy_record, dict):
        target_head = target[:40]
        if len(target_head) >= 10:
            for r in candidate_rows:
                value = normalize_copy_text(ws.cell(row=r, column=copy_col).value)
                if value[:40] == target_head:
                    matched_row = r
                    break

    if matched_row is None:
        wlog("提醒：未在 Excel 中找到匹配的已用文案，未删除。")
        wb.close()
        return False

    try:
        ws.delete_rows(matched_row, 1)
        wb.save(p)
        wb.close()
        wlog(f"已删除 Excel 中已用文案：第 {matched_row} 行。")
        return True
    except PermissionError:
        wb.close()
        wlog(f"删除文案失败：Excel 文件被占用，无法保存：{p}。请关闭 WPS/Excel 里的这个文件。")
        return False
    except Exception as e:
        wb.close()
        wlog(f"删除文案失败：保存 Excel 时出错：{repr(e)}")
        return False


def save_debug(cfg, page, name):
    try:
        d = Path(cfg["debug_dir"])
        d.mkdir(parents=True, exist_ok=True)
        path = d / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{name}.png"
        page.screenshot(path=str(path), full_page=True)
        wlog(f"调试截图：{path}")
    except Exception as e:
        wlog(f"截图失败：{e}")


def has_text(page, text, timeout=500):
    try:
        return page.get_by_text(text, exact=False).first.is_visible(timeout=timeout)
    except Exception:
        return False


def has_any(page, texts, timeout=500):
    return any(has_text(page, t, timeout) for t in texts)


def click_text(page, texts, timeout=3500):
    for t in texts:
        try:
            loc = page.get_by_text(t, exact=False).first
            loc.wait_for(state="visible", timeout=timeout)
            loc.click(force=True, timeout=timeout)
            wlog(f"点击：{t}")
            step_wait(reason=f"点击 {t} 后等待")
            return True
        except Exception:
            pass
    for t in texts:
        try:
            loc = page.locator("button").filter(has_text=t).first
            loc.wait_for(state="visible", timeout=1500)
            loc.click(force=True, timeout=timeout)
            wlog(f"点击按钮：{t}")
            step_wait(reason=f"点击按钮 {t} 后等待")
            return True
        except Exception:
            pass
    return False


def first_visible(page, selector):
    try:
        loc = page.locator(selector)
        for i in range(min(loc.count(), 100)):
            item = loc.nth(i)
            try:
                if item.is_visible(timeout=250):
                    return item
            except Exception:
                pass
    except Exception:
        pass
    return None


def visible_count(page, selector):
    try:
        loc = page.locator(selector)
        n = 0
        for i in range(min(loc.count(), 150)):
            try:
                if loc.nth(i).is_visible(timeout=180):
                    n += 1
            except Exception:
                pass
        return n
    except Exception:
        return 0


class AccountLoggedOutError(RuntimeError):
    """浏览器账号已明确进入登录或安全验证页面。"""


class DailyPublishLimitError(RuntimeError):
    """抖音明确提示今天投稿次数已达到上限。"""


DAILY_PUBLISH_LIMIT_MARKERS = (
    "抱歉，今天投稿次数已达到上限，请明天再试",
    "今天投稿次数已达到上限，请明天再试",
    "今天投稿次数已达到上限",
    "今日投稿次数已达到上限",
    "投稿次数已达到上限",
)


def text_indicates_daily_publish_limit(text):
    """兼容弹窗中的换行/空格差异，识别今日投稿次数上限提示。"""
    compact = re.sub(r"\s+", "", str(text or ""))
    return any(re.sub(r"\s+", "", marker) in compact for marker in DAILY_PUBLISH_LIMIT_MARKERS)


def daily_publish_limit_visible(page):
    try:
        body = page.locator("body").inner_text(timeout=1200)
    except Exception:
        body = ""
    return text_indicates_daily_publish_limit(body)


def normalize_retry_count(value, default=0):
    try:
        return max(0, min(99, int(value)))
    except Exception:
        return max(0, min(99, int(default or 0)))


class PublishStepFailedError(RuntimeError):
    """某个发布步骤耗尽前端配置的单步骤重试次数。"""

    def __init__(self, step_name, attempts, cause):
        self.step_name = str(step_name or "未知步骤")
        self.attempts = max(1, int(attempts or 1))
        self.cause = cause
        super().__init__(f"步骤“{self.step_name}”连续 {self.attempts} 次未完成：{cause}")


class PublishWorkflowFailedError(RuntimeError):
    """当前作品耗尽前端配置的总体流程重试次数。"""

    def __init__(self, attempts, cause):
        self.attempts = max(1, int(attempts or 1))
        self.cause = cause
        self.step_name = getattr(cause, "step_name", "未知步骤")
        super().__init__(f"当前作品总体流程连续 {self.attempts} 次未完成，最后失败步骤为“{self.step_name}”：{cause}")


def concise_publish_failure(exc, limit=240):
    """提取最后失败步骤和最底层原因，供账号队列运行状态与状态文件共同使用。"""
    step_name = str(getattr(exc, "step_name", "未知步骤") or "未知步骤")
    cause = exc
    visited = set()
    while getattr(cause, "cause", None) is not None and id(cause) not in visited:
        visited.add(id(cause))
        cause = cause.cause
    reason = str(cause or exc).strip() or repr(exc)
    text = f"{step_name}失败：{reason}"
    return text if len(text) <= limit else text[: max(1, limit - 1)] + "…"


def run_publish_step(cfg, page, step_name, action):
    """只重试当前失败步骤，不回到图片上传或重新输入之前已完成的步骤。"""
    retry_count = normalize_retry_count(cfg.get("retry_times", 3), 3)
    total_attempts = retry_count + 1
    for attempt in range(1, total_attempts + 1):
        try:
            result = action()
            if attempt > 1:
                wlog(f"步骤“{step_name}”第 {attempt}/{total_attempts} 次执行成功。")
            return result
        except (AccountLoggedOutError, DailyPublishLimitError):
            raise
        except Exception as exc:
            wlog(f"步骤“{step_name}”第 {attempt}/{total_attempts} 次失败：{repr(exc)}")
            if page is not None:
                safe_name = re.sub(r"[^0-9A-Za-z_-]+", "_", str(step_name)).strip("_") or "publish_step"
                save_debug(cfg, page, f"step_{safe_name}_{attempt}_failed")
            if attempt >= total_attempts:
                raise PublishStepFailedError(step_name, total_attempts, exc) from exc
            step_wait(cfg, f"步骤“{step_name}”失败，准备只重试本步骤")
    raise AssertionError("单步骤重试循环异常结束")


def run_publish_workflow(cfg, workflow, before_overall_retry=None):
    """单步骤重试耗尽后，按前端配置从上传开始总体重跑当前作品。"""
    overall_retry_count = normalize_retry_count(cfg.get("overall_retry_times", 1), 1)
    total_attempts = overall_retry_count + 1
    last_error = None
    for attempt in range(1, total_attempts + 1):
        try:
            return workflow(attempt, total_attempts)
        except (AccountLoggedOutError, DailyPublishLimitError):
            raise
        except PublishStepFailedError as exc:
            last_error = exc
            if attempt >= total_attempts:
                raise PublishWorkflowFailedError(total_attempts, exc) from exc
            wlog(
                f"总体流程第 {attempt}/{total_attempts} 次在步骤“{exc.step_name}”耗尽单步骤重试；"
                "准备重新进入发布页并从图片上传开始重跑当前作品。"
            )
            if before_overall_retry is not None:
                before_overall_retry(attempt, exc)
            step_wait(cfg, "当前作品总体重试：准备重新上传")
    raise PublishWorkflowFailedError(total_attempts, last_error or RuntimeError("未知发布错误"))


def creator_login_required(page):
    """只在强登录标识成立时判定掉号，普通网络失败不误判为掉号。"""
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""
    if any(marker in url for marker in ("passport.douyin.com", "/login", "sso/login", "verifycenter")):
        return True
    try:
        body = page.locator("body").inner_text(timeout=2500)
    except Exception:
        body = ""
    strong_markers = [
        "扫码登录", "手机号登录", "验证码登录", "账号密码登录",
        "请使用抖音扫码登录", "登录后即可发布", "安全验证", "身份验证",
    ]
    return any(marker in body for marker in strong_markers)


def wait_login(page):
    """兼容旧调用：新版不等待控制台输入，确认掉号后交给任务编排器跳号。"""
    if creator_login_required(page):
        raise AccountLoggedOutError("检测到登录或安全验证页面，当前浏览器账号可能已掉号。")
    return False


def image_upload_completed(page):
    """
    V21：严格判断图片是否真正上传完成。
    注意：不能把"封面设置 / 编辑图片 / 选择一张图片作为封面 / 基础信息"当作上传完成，
    因为这些文字在未上传图片时也会出现。

    真正完成标识：
    1. 已添加1张图片 / 已添加 1 张图片 / 已添加N张图片
    2. 清空并重新上传
    3. 编辑图片区块里出现已添加图片数量
    """
    try:
        body = page.locator("body").inner_text(timeout=2500)
    except Exception:
        body = ""

    if any(w in body for w in ["上传失败", "图片处理失败", "重新上传失败"]):
        return False

    # 明确的上传完成文字
    if re.search(r"已添加\s*\d+\s*张图片", body):
        return True

    if "清空并重新上传" in body:
        return True

    # DOM 层面：只在编辑图片/封面相关区域里确认"已添加N张图片"
    try:
        ok = page.evaluate("""
        () => {
          function visible(el){
            const r=el.getBoundingClientRect();
            const s=getComputedStyle(el);
            return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden';
          }
          function txt(el){ return ((el.innerText||el.textContent||'')+'').trim(); }

          const bodyText = txt(document.body);
          if (bodyText.includes('上传失败') || bodyText.includes('图片处理失败')) return false;

          const nodes=[...document.querySelectorAll('div,section,span,p,button')].filter(visible);
          for (const el of nodes) {
            const t = txt(el);
            if (!t || t.length > 300) continue;
            if (/已添加\\s*\\d+\\s*张图片/.test(t)) return true;
            if (t.includes('清空并重新上传')) return true;
          }

          return false;
        }
        """)
        return bool(ok)
    except Exception:
        return False


def wait_uploaded_with_config(cfg, page):
    interval = float(cfg.get("upload_check_interval_seconds", 3) or 3)
    max_wait = float(cfg.get("upload_max_wait_seconds", 180) or 180)
    interval = max(1, interval)
    max_wait = max(interval, max_wait)
    wlog(f"判断图片是否上传完成，检测间隔={interval}秒，最大等待={max_wait}秒。")
    start = time.time()
    while True:
        if image_upload_completed(page):
            wlog("图片上传完成：检测到编辑图片/已添加图片/封面设置等完成标识。")
            return True
        if has_any(page, ["上传失败", "图片处理失败", "重新上传"], 800):
            raise RuntimeError("检测到图片上传失败。")
        if time.time() - start >= max_wait:
            raise RuntimeError(f"等待图片上传超过 {max_wait} 秒，仍未检测到完成标识。")
        wlog(f"暂未检测到图片上传完成，等待 {interval:.1f} 秒后重试。")
        countdown_sleep(ACTIVE_CFG or {}, interval, "检测图片上传状态", "upload_check")


def validate_uploaded_image(cfg, page):
    ok = image_upload_completed(page)
    if not ok:
        save_debug(cfg, page, "validate_uploaded_image_failed")
        raise RuntimeError("图片上传状态未完成，未检测到编辑图片/已添加图片/封面设置。")
    wlog("校验通过：图片已上传完成。")


def split_text_topics(text):
    """
    从文案里提取话题，并从正文中彻底清除话题残留。
    修复点：
    1. 支持 #话题、##话题、###话题、＃话题 等写法。
    2. 删除正文里因为双 # / 多 # 导致残留的 ###。
    3. 返回去重后的干净话题，不带 #。
    """
    text = str(text or "").replace("＃", "#")

    # 支持一个或多个 #，也支持 # 后面有空格的情况
    topics = re.findall(r"#+\s*([\u4e00-\u9fa5A-Za-z0-9_]+)", text)
    seen = []
    for t in topics:
        t = str(t).strip().lstrip("#").strip()
        if t and t not in seen:
            seen.append(t)

    # 删除完整话题，避免正文里留下单独的 #
    body = re.sub(r"#+\s*[\u4e00-\u9fa5A-Za-z0-9_]+", "", text)
    # 兜底删除所有孤立/残留 #，解决页面出现 ### 的问题
    body = re.sub(r"[#＃]+", "", body)
    body = re.sub(r"[ \t]{2,}", " ", body)
    body = re.sub(r"\n{3,}", "\n\n", body).strip()
    return body, seen


def upload_tab_ready(page):
    """
    V24：判断是否已经在"发布图文"上传页。
    典型页面：creator.douyin.com/creator-micro/content/upload?default-tab=3
    顶部有"发布视频 / 发布图文 / 发布全景视频 / 发布文章"，中间有红色"上传图文"。
    """
    try:
        url = page.url.lower()
    except Exception:
        url = ""
    try:
        text = page.locator("body").inner_text(timeout=2500)
    except Exception:
        text = ""
    if "content/upload" in url and ("发布图文" in text or "上传图文" in text):
        return True
    if "上传图文" in text and "图片格式" in text and "图片大小" in text:
        return True
    if "点击上传" in text and "直接将图片文件拖入此区域" in text:
        return True
    return False


def click_upload_graphic_button_or_input(cfg, page, image_path):
    """
    V24：在发布图文上传页，直接上传图片。
    不处理"继续编辑 / 放弃"，也不重复点击发布入口。
    """
    wlog("当前为发布图文上传页，直接点击上传图文/文件 input 上传图片。")
    try:
        inp = page.locator("input[type='file']")
        chosen = None
        for i in range(min(inp.count(), 30)):
            item = inp.nth(i)
            try:
                accept = (item.get_attribute("accept") or "").lower()
                if "image" in accept or ".jpg" in accept or ".png" in accept or ".jpeg" in accept or ".webp" in accept:
                    chosen = item
                    break
            except Exception:
                pass
        if chosen is None and inp.count() > 0:
            chosen = inp.first
        if chosen is not None:
            chosen.set_input_files(str(image_path), timeout=30000)
            wlog("已通过文件 input 上传图片。")
            return True
    except Exception as e:
        wlog(f"直接设置上传 input 失败，改点上传图文按钮：{repr(e)}")

    upload_words = ["上传图文", "点击上传", "上传图片", "选择图片", "选择文件", "上传"]
    for word in upload_words:
        try:
            with page.expect_file_chooser(timeout=8000) as fc:
                ok = page.evaluate("""
                (word) => {
                  function visible(el){
                    const r = el.getBoundingClientRect();
                    const s = getComputedStyle(el);
                    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
                  }
                  function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
                  const nodes = [...document.querySelectorAll('button,[role=button],span,div,a')].filter(visible);
                  const arr = [];
                  for (const el of nodes) {
                    const t = txt(el);
                    if (!t || t.length > 80) continue;
                    if (!t.includes(word)) continue;
                    const r = el.getBoundingClientRect();
                    let score = 0;
                    if (t === word) score += 50;
                    if (word === '上传图文') score += 100;
                    if (r.top > 150 && r.top < window.innerHeight * 0.8) score += 20;
                    arr.push({el, score});
                  }
                  arr.sort((a,b)=>b.score-a.score);
                  if (arr.length) {
                    (arr[0].el.closest('button') || arr[0].el.closest('[role=button]') || arr[0].el).click();
                    return true;
                  }
                  return false;
                }
                """, word)
                if not ok:
                    raise RuntimeError(f"没找到按钮：{word}")
            fc.value.set_files(str(image_path))
            wlog(f"已点击'{word}'并选择图片。")
            return True
        except Exception:
            pass
    return False

def platform_image_publish_page_ready(page):
    """
    V24：图文发布准备状态包括：
    1. 图文编辑页 post/image
    2. 图文上传页 content/upload?default-tab=3
    """
    if upload_tab_ready(page):
        return True
    try:
        url = page.url.lower()
    except Exception:
        url = ""
    if "post/image" in url or "media_type=image" in url:
        return True
    try:
        body = page.locator("body").inner_text(timeout=2000)
    except Exception:
        body = ""
    return ("作品描述" in body and "封面设置" in body and "选择音乐" in body)


def click_creator_publish_entry(cfg, page):
    """
    V29：不点击会触发新标签的发布入口。
    所有页面跳转都在当前标签里用 page.goto 完成。
    """
    wait_sec = float(cfg.get("creator_center_wait_seconds", 8) or 8)
    if wait_sec < 0:
        wait_sec = 0

    wlog("当前标签打开抖音创作者中心首页。")
    page.goto(cfg.get("creator_url", "https://creator.douyin.com/"), wait_until="domcontentloaded", timeout=60000)
    safe_bring_to_front(cfg, page)

    wait_login(page)

    upload_url = "https://creator.douyin.com/creator-micro/content/upload?default-tab=3"
    target_url = ""
    deadline = time.time() + max(1.0, wait_sec)
    while time.time() < deadline and not target_url:
        candidates = [
            page.locator('a[href*="/content/upload"],a[href*="/content/post"]').filter(
                has_text=re.compile(r"发布|图文")
            ),
            page.get_by_role("link", name=re.compile(r"发布作品|发布图文|高清发布|发布")),
            page.locator('[data-e2e*="publish" i],[data-testid*="publish" i]').filter(
                has_text=re.compile(r"发布|图文")
            ),
        ]
        for locator in candidates:
            try:
                for index in range(min(locator.count(), 12)):
                    item = locator.nth(index)
                    if not item.is_visible(timeout=200):
                        continue
                    href = str(item.get_attribute("href", timeout=500) or "").strip()
                    if not href or "javascript:" in href.lower():
                        continue
                    resolved = urllib.parse.urljoin(page.url, href)
                    if "/content/upload" in resolved or "/content/post" in resolved:
                        target_url = resolved
                        break
                if target_url:
                    break
            except Exception:
                pass
        if not target_url:
            page.wait_for_timeout(300)

    if target_url:
        wlog("已通过发布入口的 href/role/data 属性定位图文发布模块，并在当前标签打开。")
        if "/content/post" in target_url and not any(marker in target_url for marker in ("/image", "media_type=image", "type=image")):
            target_url = upload_url
        if "/content/upload" in target_url and "default-tab=" not in target_url:
            separator = "&" if "?" in target_url else "?"
            target_url += separator + "default-tab=3"
    else:
        target_url = upload_url
        wlog("未读到可用发布入口 href，使用图文上传页地址作为无坐标兜底。")

    page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
    step_wait(cfg, "当前标签打开发布图文上传页后等待")

    if platform_image_publish_page_ready(page):
        wlog("已在当前标签进入发布图文页面。")
        return True

    if creator_login_required(page):
        save_debug(cfg, page, "browser_account_logged_out")
        raise AccountLoggedOutError("当前浏览器未能进入图文发布页，并检测到登录/验证页面。")

    save_debug(cfg, page, "image_publish_page_not_ready")
    raise RuntimeError("没有成功进入图文发布页面；未检测到明确登录页，按临时页面故障处理。")



def enter_publish_page(cfg, page):
    """
    V22：每次都从创作者中心首页重新进入发布图文页。
    不再保留浏览器原网页直接继续。
    """
    click_creator_publish_entry(cfg, page)
    wait_login(page)


def upload_image(cfg, page, image_path):
    """
    V24：如果当前已经是发布图文上传页，直接点击"上传图文"上传。
    不管"继续编辑 / 放弃"，不重复点击发布入口。
    """
    wlog(f"上传图片：{Path(image_path).name}")
    if upload_tab_ready(page):
        ok = click_upload_graphic_button_or_input(cfg, page, image_path)
        if ok:
            return
        save_debug(cfg, page, "upload_graphic_button_failed")
        raise RuntimeError("当前是发布图文上传页，但没能点击上传图文按钮或上传 input。")

    try:
        inp = page.locator("input[type='file']")
        chosen = None
        for i in range(min(inp.count(), 30)):
            item = inp.nth(i)
            try:
                accept = (item.get_attribute("accept") or "").lower()
                if "image" in accept or ".jpg" in accept or ".png" in accept or ".jpeg" in accept or ".webp" in accept:
                    chosen = item
                    break
            except Exception:
                pass
        if chosen is None and inp.count() > 0:
            chosen = inp.first
        if chosen is not None:
            chosen.set_input_files(str(image_path), timeout=30000)
            wlog("已选择图片。")
            return
    except Exception as e:
        wlog(f"直接设置图片 input 失败，改用文件选择器：{repr(e)}")

    for t in ["上传图文", "点击上传", "上传图片", "选择图片", "添加图片", "选择文件", "上传"]:
        try:
            with page.expect_file_chooser(timeout=8000) as fc:
                click_text(page, [t], timeout=5000)
            fc.value.set_files(str(image_path))
            wlog("已通过文件选择器选择图片。")
            return
        except Exception:
            pass
    save_debug(cfg, page, "upload_entry_not_found")
    raise RuntimeError("没有找到图片上传入口。")


def wait_after_upload_to_editor(cfg, page):
    """
    V24：点击上传图文/选择图片后，等待页面进入图文编辑页。
    """
    max_wait = float(cfg.get("upload_max_wait_seconds", 180) or 180)
    interval = float(cfg.get("upload_check_interval_seconds", 3) or 3)
    interval = max(1, interval)
    start_time = time.time()
    while True:
        try:
            text = page.locator("body").inner_text(timeout=2500)
        except Exception:
            text = ""
        if ("作品描述" in text or "基础信息" in text) and ("选择音乐" in text or "封面设置" in text):
            wlog("已进入图文编辑页。")
            return True
        if re.search(r"已添加\s*\d+\s*张图片", text):
            wlog("检测到已添加图片，继续。")
            return True
        if "上传失败" in text or "图片处理失败" in text:
            raise RuntimeError("图片上传失败。")
        if time.time() - start_time > max_wait:
            save_debug(cfg, page, "wait_editor_after_upload_timeout")
            raise RuntimeError(f"上传后等待进入编辑页超过 {max_wait} 秒。")
        wlog(f"等待进入图文编辑页，{interval:.1f} 秒后重试。")
        countdown_sleep(ACTIVE_CFG or {}, interval, "等待进入图文编辑页", "edit_page_check")

def editor_descriptor(item):
    try:
        return item.evaluate("""
        el => [
          el.getAttribute('placeholder') || '',
          el.getAttribute('data-placeholder') || '',
          el.getAttribute('aria-label') || '',
          el.getAttribute('name') || '',
          el.closest('label')?.innerText || '',
          el.parentElement?.innerText?.slice(0, 80) || ''
        ].join(' ')
        """)
    except Exception:
        return ""


def find_editor(page):
    preferred = [
        "[contenteditable='true'][data-placeholder*='作品描述']",
        "[contenteditable='true'][aria-label*='作品描述']",
        "textarea[placeholder*='作品描述']",
        "textarea[placeholder*='描述']",
        "[contenteditable='true'][data-placeholder*='正文']",
    ]
    for sel in preferred:
        item = first_visible(page, sel)
        if item:
            return item

    selectors = ["div[contenteditable='true']", "[contenteditable='true']", "textarea", "[role='textbox']"]
    for sel in selectors:
        try:
            loc = page.locator(sel)
            for index in range(min(loc.count(), 100)):
                item = loc.nth(index)
                if not item.is_visible(timeout=250):
                    continue
                if "标题" in editor_descriptor(item):
                    continue
                return item
        except Exception:
            pass
    return None


def find_title_editor(page):
    selectors = [
        "input[placeholder*='标题']",
        "textarea[placeholder*='标题']",
        "[contenteditable='true'][data-placeholder*='标题']",
        "[contenteditable='true'][aria-label*='标题']",
        "input[aria-label*='标题']",
        "textarea[aria-label*='标题']",
        "[role='textbox'][aria-label*='标题']",
    ]
    for sel in selectors:
        item = first_visible(page, sel)
        if item:
            return item
    return None


def get_input_text(item):
    try:
        return item.evaluate("el => ('value' in el ? el.value : (el.innerText || el.textContent || ''))")
    except Exception:
        return ""


def fill_title(cfg, page, title):
    title = copy_cell_text(title)
    if not title:
        return False
    editor = find_title_editor(page)
    if not editor:
        save_debug(cfg, page, "title_editor_not_found")
        raise RuntimeError("Excel 已填写标题，但发布页没有找到标题输入框。")
    editor.click(force=True)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.type(title, delay=0)
    step_wait(cfg, "输入标题后等待")
    return True


def validate_title_filled(cfg, page, title):
    title = copy_cell_text(title)
    if not title:
        return True
    expected = normalize_text_for_check(title)
    for attempt in range(1, 4):
        editor = find_title_editor(page)
        actual = normalize_text_for_check(get_input_text(editor)) if editor else ""
        if expected and expected in actual:
            wlog(f"校验通过：标题已写入，共 {len(expected)} 个有效字符。")
            return True
        wlog(f"第 {attempt}/3 次标题校验暂未通过，等待输入框状态稳定。")
        time.sleep(0.8)
    save_debug(cfg, page, "validate_title_filled_failed")
    raise RuntimeError("标题校验失败：标题输入框中没有检测到对应内容。")


def fill_copy(cfg, page, text):
    """
    填写正文并点亮话题（逐字输入版，不占用剪贴板）。
    修复点：
    1. 话题每输入一次后等待识别，避免太快导致话题没点亮。
    2. 函数缩进修正，可直接运行。
    3. 使用 page.keyboard.type() 逐字输入，不占用剪贴板。
    """
    body, topics = split_text_topics(text)
    editor = find_editor(page)
    if not editor:
        save_debug(cfg, page, "editor_not_found")
        raise RuntimeError("没有找到文案输入框。")

    editor.click(force=True)
    step_wait(cfg, "点击文案框后等待")

    # 先清空文案框
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    step_wait(cfg, "清空文案框后等待")

    # 逐字输入正文（不占用剪贴板）
    input_text = body if body else text
    wlog(f"逐字输入正文，共 {len(input_text)} 字")
    page.keyboard.type(input_text, delay=0)
    step_wait(cfg, "输入文案后等待")

    if topics:
        wlog("点亮话题：" + "、".join(topics))
        page.keyboard.press("End")
        page.keyboard.press("Enter")

        # 等待抖音话题输入区域稳定
        countdown_sleep(cfg, float(cfg.get("topic_first_wait_seconds", 1.5) or 1.5), "话题输入区域稳定", "topic_wait")

        # 每个话题单独输入，输入后等待平台识别，再按 Enter 点亮
        topic_wait = float(cfg.get("topic_wait_seconds", 2.5) or 2.5)
        topic_after_enter_wait = float(cfg.get("topic_after_enter_wait_seconds", 1) or 1)

        for topic in topics:
            topic = str(topic).strip().replace("＃", "#").lstrip("#").strip()
            if not topic:
                continue
            # 逐字输入话题 #tag（不占用剪贴板）
            page.keyboard.type("#" + topic, delay=0)

            wlog(f"等待识别话题：#{topic}")
            countdown_sleep(cfg, topic_wait, f"识别话题：#{topic}", "topic_wait")

            page.keyboard.press("Enter")
            countdown_sleep(cfg, topic_after_enter_wait, f"确认话题：#{topic}", "topic_wait")
            page.keyboard.type(" ")


def normalize_text_for_check(value):
    """
    文案校验专用清理：
    忽略换行、空格、标点以及页面可能插入的零宽字符。
    """
    value = str(value or "")
    for ch in ("\u200b", "\u200c", "\u200d", "\ufeff"):
        value = value.replace(ch, "")
    return re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9]+", "", value)


def get_editor_text_for_check(page):
    """
    只读取可见的正文输入区域，优先取文字最长的输入框，
    避免右侧快速检测、页面换行或整页动态刷新造成误判。
    """
    try:
        return page.evaluate("""
        () => {
          function visible(el) {
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 &&
                   s.display !== 'none' && s.visibility !== 'hidden';
          }

          const selectors = [
            "div[contenteditable='true']",
            "[contenteditable='true']",
            "textarea",
            "[role='textbox']"
          ];

          const elements = [...document.querySelectorAll(selectors.join(","))]
            .filter(visible)
            .filter(el => {
              const descriptor = [
                el.getAttribute('placeholder') || '',
                el.getAttribute('data-placeholder') || '',
                el.getAttribute('aria-label') || '',
                el.getAttribute('name') || '',
                el.closest('label')?.innerText || '',
                el.parentElement?.innerText?.slice(0, 80) || ''
              ].join(' ');
              return !descriptor.includes('标题');
            });

          const texts = elements.map(el => {
            if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
              return el.value || "";
            }
            return el.innerText || el.textContent || "";
          }).filter(Boolean);

          texts.sort((a, b) => b.length - a.length);
          return texts[0] || "";
        }
        """)
    except Exception:
        return ""


def validate_copy_filled(cfg, page, text):
    """
    校验文案是否真正写入。

    修复旧版误判：
    1. 不再只读取整页 body 做严格原文匹配；
    2. 忽略换行、标点、空格和零宽字符；
    3. 等待抖音快速检测与话题识别稳定后重复校验；
    4. 编辑框读取失败时，再使用整页文本作为兜底。
    """
    body, _ = split_text_topics(text)
    expected = normalize_text_for_check(body or text)

    if not expected:
        raise RuntimeError("原始文案为空，无法进行文案校验。")

    actual = ""
    for attempt in range(1, 6):
        editor_text = get_editor_text_for_check(page)
        actual = normalize_text_for_check(editor_text)

        # 编辑框暂时读取不到时，才使用整页文本兜底。
        if not actual:
            try:
                page_text = page.locator("body").inner_text(timeout=5000)
            except Exception:
                page_text = ""
            actual = normalize_text_for_check(page_text)

        if len(expected) < 12:
            ok = expected in actual
        else:
            size = 12
            middle = max(0, len(expected) // 2 - size // 2)
            probes = [
                expected[:size],
                expected[middle:middle + size],
                expected[-size:],
            ]
            hits = [probe for probe in probes if probe and probe in actual]
            # 开头命中即可；若开头有少量差异，则中间和结尾同时命中也算成功。
            ok = probes[0] in actual or len(hits) >= 2

        if ok:
            wlog(
                f"校验通过：文案已写入，"
                f"读取到 {len(actual)} 个有效字符。"
            )
            return True

        wlog(
            f"第 {attempt}/5 次文案校验暂未通过，"
            f"等待快速检测和输入框状态稳定。"
        )
        time.sleep(1.2)

    save_debug(cfg, page, "validate_copy_filled_failed")
    wlog(f"预期文案片段：{expected[:36]}")
    wlog(f"实际输入框片段：{actual[:36]}")
    raise RuntimeError("文案校验失败：输入框中没有检测到对应文案内容。")


def _legacy_verify_music_selected(page):
    """
    V20：严格判断音乐是否真的使用成功。
    只看"选择音乐"这一行/扩展信息区域，不能用页面其他地方的 00:xx 时长误判。

    成功条件：
    1. 选择音乐区域出现"修改音乐 / 更换音乐 / 删除音乐"
    2. 或选择音乐区域里没有"点击添加合适作品风格音乐"，且有歌曲名 + 时长

    失败条件：
    选择音乐区域仍然出现"点击添加合适作品风格音乐"
    """
    try:
        result = page.evaluate("""
        () => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }

          const nodes = [...document.querySelectorAll('div,section,li')].filter(visible);
          const rows = [];

          for (const el of nodes) {
            const t = txt(el);
            if (!t || t.length > 600) continue;
            // 只抓"选择音乐"相关的那一块，避免抓到右侧预览或整页文本
            if (t.includes('选择音乐') || t.includes('点击添加合适作品风格音乐') || t.includes('修改音乐') || t.includes('更换音乐')) {
              const r = el.getBoundingClientRect();
              // 只考虑页面左侧/中间编辑区，不考虑右侧预览手机
              if (r.left > window.innerWidth * 0.75) continue;
              rows.push({
                text: t,
                x: r.left, y: r.top, w: r.width, h: r.height,
                area: r.width * r.height
              });
            }
          }

          // 优先选择面积适中的行块，避免选中整页大容器
          rows.sort((a,b) => {
            const as = (a.text.includes('选择音乐') ? 1000 : 0) - Math.abs(a.area - 60000);
            const bs = (b.text.includes('选择音乐') ? 1000 : 0) - Math.abs(b.area - 60000);
            return bs - as;
          });

          const musicBlock = rows[0] || null;
          const text = musicBlock ? musicBlock.text : '';

          const hasAddPrompt = text.includes('点击添加合适作品风格音乐');
          const hasModify = text.includes('修改音乐') || text.includes('更换音乐') || text.includes('删除音乐');
          const hasDuration = /\\d{1,2}:\\d{2}/.test(text);

          return {
            ok: !hasAddPrompt && (hasModify || hasDuration),
            hasAddPrompt,
            hasModify,
            hasDuration,
            text: text.slice(0, 300),
            rowsCount: rows.length
          };
        }
        """)
    except Exception as e:
        wlog(f"音乐校验异常：{repr(e)}")
        return False

    wlog("音乐校验详情：" + json.dumps(result, ensure_ascii=False)[:600])

    if result.get("hasAddPrompt"):
        wlog("音乐校验失败：选择音乐区域仍显示'点击添加合适作品风格音乐'。")
        return False

    if result.get("ok"):
        wlog("音乐校验通过：选择音乐区域已显示已使用音乐。")
        return True

    return False


def _legacy_validate_music_added(cfg, page):
    if not verify_music_selected(page):
        save_debug(cfg, page, "validate_music_failed")
        raise RuntimeError("音乐未使用成功：选择音乐区域仍未显示已使用音乐。")
    wlog("校验通过：音乐已选择。")


def music_selection_state(page):
    """读取主编辑区的音乐模块，返回 selected / unselected / unknown 三态。"""
    try:
        return page.evaluate(r"""
        () => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' &&
                   s.visibility !== 'hidden' && Number(s.opacity || 1) > 0;
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').replace(/\s+/g, ' ').trim(); }
          function insideMusicPopup(el){
            const popup = el.closest('[role="dialog"],[aria-modal="true"],[data-e2e*="music" i],[data-testid*="music" i]');
            if (!popup) return false;
            const t = txt(popup);
            return t.includes('选择音乐') && (t.includes('使用') || t.includes('热门') || t.includes('推荐'));
          }

          const anchors = [...document.querySelectorAll('div,section,li,label,span,button')]
            .filter(visible)
            .filter(el => {
              const t = txt(el);
              return t === '选择音乐' || t.includes('点击添加合适作品风格音乐') ||
                     t.includes('修改音乐') || t.includes('更换音乐') || t.includes('删除音乐');
            });
          const candidates = [];
          for (const anchor of anchors) {
            if (insideMusicPopup(anchor)) continue;
            let el = anchor;
            for (let depth = 0; el && depth < 6; depth++, el = el.parentElement) {
              if (!visible(el)) continue;
              const text = txt(el);
              const r = el.getBoundingClientRect();
              const area = r.width * r.height;
              if (!text || text.length > 500 || area < 1200 || area > 500000) continue;
              if (!(text.includes('选择音乐') || text.includes('点击添加合适作品风格音乐') ||
                    text.includes('修改音乐') || text.includes('更换音乐') || text.includes('删除音乐'))) continue;
              if (text.includes('热门榜') || text.includes('搜索音乐') || text.includes('推荐音乐')) continue;
              const explicitEmpty = text.includes('点击添加合适作品风格音乐');
              const hasAction = /修改音乐|更换音乐|删除音乐|已选择|已使用/.test(text);
              const hasDuration = /\d{1,2}:\d{2}/.test(text);
              const hasImage = !!el.querySelector('img');
              let score = 0;
              if (explicitEmpty) score += 500;
              if (hasAction) score += 600;
              if (hasDuration) score += 250;
              if (hasImage) score += 80;
              if (text.includes('选择音乐')) score += 80;
              score -= Math.abs(area - 50000) / 3000;
              candidates.push({text, explicitEmpty, hasAction, hasDuration, hasImage, score});
            }
          }
          candidates.sort((a,b) => b.score - a.score);
          const block = candidates[0] || null;
          if (!block) return {state:'unknown', selected:false, explicitEmpty:false, text:'', candidates:0};
          if (block.explicitEmpty) {
            return {state:'unselected', selected:false, explicitEmpty:true, text:block.text.slice(0,300), candidates:candidates.length};
          }
          const selected = block.hasAction || block.hasDuration ||
            (block.hasImage && block.text.replace(/选择音乐|扩展信息|音乐/g, '').trim().length >= 2);
          return {
            state: selected ? 'selected' : 'unknown', selected,
            explicitEmpty:false, text:block.text.slice(0,300), candidates:candidates.length
          };
        }
        """)
    except Exception as exc:
        return {"state": "unknown", "selected": False, "explicitEmpty": False, "text": "", "error": repr(exc)}


def verify_music_selected(page, log_detail=True):
    state = music_selection_state(page)
    if log_detail:
        wlog("音乐校验详情：" + json.dumps(state, ensure_ascii=False)[:600])
    if state.get("selected"):
        if log_detail:
            wlog("音乐校验通过：音乐模块已显示歌曲信息或修改入口。")
        return True
    if state.get("explicitEmpty") and log_detail:
        wlog("音乐尚未选择：音乐模块仍显示添加音乐提示。")
    return False


def wait_music_selection_applied(page, timeout_seconds=10):
    """点击使用后短轮询音乐模块；页面响应慢时不立即误判失败。"""
    deadline = time.time() + max(1.0, float(timeout_seconds or 10))
    last_state = {"state": "unknown", "selected": False, "explicitEmpty": False}
    panel_closed_observations = 0
    while time.time() < deadline:
        last_state = music_selection_state(page)
        if last_state.get("selected"):
            wlog("点击使用后已检测到音乐生效。")
            return True, last_state
        if not platform_music_panel_opened(page):
            panel_closed_observations += 1
            if panel_closed_observations == 1:
                wlog("音乐面板已关闭，继续等待主音乐模块刷新。")
        page.wait_for_timeout(400)
    return False, last_state


def validate_music_added(cfg, page):
    ok, state = wait_music_selection_applied(page, timeout_seconds=6)
    if not ok:
        save_debug(cfg, page, "music_click_state_not_applied")
        if state.get("explicitEmpty"):
            raise RuntimeError("点击使用后音乐未生效：音乐模块仍明确显示未选择。")
        raise RuntimeError("点击使用后无法确认音乐状态：音乐模块未显示歌曲信息或修改入口。")
    wlog("校验通过：音乐已选择。")


def _legacy_platform_music_panel_opened(page):
    """
    严格判断右侧平台音乐面板是否打开。
    不能因为页面本身有"选择音乐"字段就误判。
    """
    try:
        return bool(page.evaluate("""
        () => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
          const vw = window.innerWidth;
          const nodes = [...document.querySelectorAll('div,section,[role=dialog]')].filter(visible);
          for (const el of nodes) {
            const r = el.getBoundingClientRect();
            const t = txt(el);
            // 音乐面板在右侧抽屉区
            if (r.left < vw * 0.45) continue;
            if (!t.includes('选择音乐')) continue;
            if (t.includes('热门榜') || t.includes('热门音乐') || t.includes('搜索音乐') || t.includes('原创榜') || t.includes('使用')) {
              return true;
            }
          }
          return false;
        }
        """))
    except Exception:
        return False


def music_panel_info(page):
    """按 role、aria/data 属性和模块内容识别音乐面板，不把屏幕坐标作为唯一条件。"""
    try:
        return page.evaluate(r"""
        () => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' &&
                   s.visibility !== 'hidden' && Number(s.opacity || 1) > 0;
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').replace(/\s+/g, ' ').trim(); }
          const selectors = [
            '[role="dialog"]','[aria-modal="true"]','[data-e2e*="music" i]',
            '[data-testid*="music" i]','aside','section','div'
          ];
          const vw = window.innerWidth, vh = window.innerHeight;
          const nodes = [...new Set(selectors.flatMap(sel => [...document.querySelectorAll(sel)]))].filter(visible);
          const candidates = [];
          for (const el of nodes) {
            const text = txt(el);
            if (!text.includes('选择音乐') || text.length > 3000) continue;
            if (!(text.includes('使用') || text.includes('热门') || text.includes('推荐') || text.includes('搜索音乐'))) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 260 || r.height < 180) continue;
            const role = el.getAttribute('role') || '';
            const data = (el.getAttribute('data-e2e') || '') + (el.getAttribute('data-testid') || '');
            const position = getComputedStyle(el).position;
            const semanticPanel = role === 'dialog' || el.getAttribute('aria-modal') === 'true' || /music/i.test(data);
            if (!semanticPanel && !['fixed','absolute','sticky'].includes(position) && r.width > vw * 0.65) continue;
            if (!semanticPanel && (r.width > vw * 0.82 || r.height > vh * 1.35 || r.top < -50)) continue;
            if (text.includes('基础信息') && text.includes('发布设置')) continue;
            let score = 0;
            if (role === 'dialog') score += 1000;
            if (el.getAttribute('aria-modal') === 'true') score += 800;
            if (/music/i.test(data)) score += 700;
            if (position === 'fixed') score += 300;
            if (r.left > vw * 0.35) score += 100;
            if (text.includes('使用')) score += 200;
            if (el.querySelector('button,[role="button"]')) score += 100;
            score += Math.min(300, el.querySelectorAll('li,[role="listitem"],img').length * 10);
            score -= Math.abs(r.width * r.height - 350000) / 10000;
            candidates.push({el, score, text, r});
          }
          candidates.sort((a,b) => b.score - a.score);
          const panel = candidates[0];
          if (!panel) return {opened:false, text:'', candidates:0};
          return {
            opened:true, text:panel.text.slice(0,500), candidates:candidates.length,
            role:panel.el.getAttribute('role') || '',
            data:(panel.el.getAttribute('data-e2e') || panel.el.getAttribute('data-testid') || ''),
            x:panel.r.left, y:panel.r.top, width:panel.r.width, height:panel.r.height
          };
        }
        """)
    except Exception as exc:
        return {"opened": False, "error": repr(exc)}


def platform_music_panel_opened(page):
    return bool(music_panel_info(page).get("opened"))


def scroll_to_music_area(cfg, page):
    """
    滚动到扩展信息里的选择音乐区域。
    """
    for i in range(1, 6):
        try:
            page.evaluate("""
            () => {
              function visible(el){
                const r = el.getBoundingClientRect();
                const s = getComputedStyle(el);
                return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
              }
              function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
              const nodes = [...document.querySelectorAll('div,section,span,label,p,button')].filter(visible);
              const el = nodes.find(n => {
                const t = txt(n);
                return t && t.length < 200 && (t.includes('点击添加合适作品风格音乐') || t.includes('选择音乐') || t.includes('扩展信息'));
              });
              if (el) el.scrollIntoView({block:'center', inline:'nearest'});
            }
            """)
        except Exception:
            pass
        step_wait(cfg, f"滚动到选择音乐区域 第{i}次")
        try:
            text = page.locator("body").inner_text(timeout=2000)
        except Exception:
            text = ""
        if "点击添加合适作品风格音乐" in text or "选择音乐" in text:
            return True
    return False


def _legacy_open_music_panel(cfg, page):
    if verify_music_selected(page):
        return True
    if platform_music_panel_opened(page):
        wlog("平台音乐面板已经打开。")
        return True

    scroll_to_music_area(cfg, page)

    # 只点击"选择音乐"区域右侧按钮或该行右侧区域
    js_click_music = """
    () => {
      function visible(el){
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
      }
      function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }

      const vw = window.innerWidth;
      const nodes = [...document.querySelectorAll('div,section,li')].filter(visible);
      const candidates = [];

      for (const el of nodes) {
        const t = txt(el);
        if (!t) continue;
        if (t.length > 800) continue;
        if (!(t.includes('选择音乐') || t.includes('点击添加合适作品风格音乐'))) continue;
        if (t.includes('热门榜') || t.includes('搜索音乐')) continue; // 排除右侧面板
        const r = el.getBoundingClientRect();
        if (r.left > vw * 0.75) continue; // 排除右侧预览/抽屉
        if (r.height < 30 || r.width < 250) continue;

        let score = 0;
        if (t.includes('点击添加合适作品风格音乐')) score += 100;
        if (t.includes('选择音乐')) score += 60;
        if (t.includes('扩展信息')) score += 10;
        // 面积适中更像音乐行
        score -= Math.abs((r.width * r.height) - 50000) / 10000;

        candidates.push({el, score, text:t.slice(0,120), x:r.left, y:r.top, w:r.width, h:r.height});
      }

      candidates.sort((a,b)=>b.score-a.score);
      const row = candidates[0];
      if (!row) return {ok:false, reason:'music row not found'};

      // 优先点击行内可见的"选择音乐"按钮/文字
      const clickables = [...row.el.querySelectorAll('button,[role=button],span,div,a')].filter(visible);
      const btns = [];
      for (const b of clickables) {
        const t = txt(b);
        const r = b.getBoundingClientRect();
        if (t === '选择音乐' || t.includes('选择音乐')) {
          btns.push({el:b, x:r.left, y:r.top, w:r.width, h:r.height, text:t});
        }
      }
      btns.sort((a,b)=>b.x-a.x);
      if (btns.length) {
        (btns[0].el.closest('button') || btns[0].el.closest('[role=button]') || btns[0].el).click();
        return {ok:true, method:'button', text:btns[0].text};
      }

      // 兜底：点击该行右侧区域，通常是"选择音乐"按钮所在位置
      const x = row.x + row.w - 70;
      const y = row.y + row.h / 2;
      const target = document.elementFromPoint(x, y);
      if (target) {
        target.click();
        return {ok:true, method:'right-area', x, y, rowText:row.text};
      }

      return {ok:false, reason:'no target at right area', rowText:row.text};
    }
    """

    for attempt in range(1, 8):
        try:
            result = page.evaluate(js_click_music)
            wlog("点击音乐入口结果：" + json.dumps(result, ensure_ascii=False)[:400])
            if result and result.get("ok"):
                step_wait(cfg, f"第 {attempt} 次点击选择音乐后等待面板")
                if platform_music_panel_opened(page):
                    wlog("平台音乐面板已打开。")
                    return True
        except Exception as e:
            wlog(f"第 {attempt} 次点击选择音乐入口失败：{repr(e)}")

        # 有些情况下点了行但没打开，再点"选择音乐"文字坐标
        try:
            rects = get_text_rects(page, "选择音乐")
            # 选择左侧/中间编辑区的文字
            rects = [r for r in rects if r.get("x", 9999) < 700]
            if rects:
                r = rects[-1]
                page.mouse.click(float(r["cx"]), float(r["cy"]))
                step_wait(cfg, "点击选择音乐文字后等待")
                if platform_music_panel_opened(page):
                    wlog("平台音乐面板已打开。")
                    return True
        except Exception:
            pass

        scroll_to_music_area(cfg, page)
        step_wait(cfg, "音乐面板未打开，准备重试")

    save_debug(cfg, page, "music_panel_not_opened")
    raise RuntimeError("没有打开平台音乐面板。")


def click_locator_with_fallback(page, locator, label, timeout=2500):
    """Playwright 语义点击 → DOM 原生 click → 真实鼠标；坐标只使用元素实际边界。"""
    try:
        locator.scroll_into_view_if_needed(timeout=timeout)
    except Exception:
        pass
    try:
        locator.click(timeout=timeout)
        wlog(f"{label}：Playwright 语义点击成功。")
        return True
    except Exception:
        pass
    try:
        locator.evaluate("el => el.click()", timeout=timeout)
        wlog(f"{label}：DOM 原生 click 成功。")
        return True
    except Exception:
        pass
    try:
        box = locator.bounding_box(timeout=timeout)
        if box:
            page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            wlog(f"{label}：真实鼠标点击元素边界中心成功。")
            return True
    except Exception:
        pass
    return False


def wait_music_panel_open(page, timeout_seconds=4):
    deadline = time.time() + max(0.5, float(timeout_seconds or 4))
    while time.time() < deadline:
        if platform_music_panel_opened(page):
            return True
        page.wait_for_timeout(250)
    return False


def open_music_panel(cfg, page):
    if verify_music_selected(page, log_detail=False):
        return True
    if platform_music_panel_opened(page):
        wlog("平台音乐面板已经打开。")
        return True

    semantic_candidates = []
    for getter in (
        lambda: page.get_by_role("button", name=re.compile(r"选择音乐|添加音乐")),
        lambda: page.get_by_text("点击添加合适作品风格音乐", exact=False),
        lambda: page.get_by_text("选择音乐", exact=True),
        lambda: page.locator('[data-e2e*="music" i],[data-testid*="music" i]').filter(has_text="选择音乐"),
    ):
        try:
            locator = getter()
            for index in range(min(locator.count(), 8)):
                item = locator.nth(index)
                if item.is_visible(timeout=300):
                    semantic_candidates.append(item)
        except Exception:
            pass

    for item in semantic_candidates:
        if click_locator_with_fallback(page, item, "打开音乐面板") and wait_music_panel_open(page, 3):
            wlog("平台音乐面板已打开。")
            return True

    # DOM 模块兜底：定位含添加提示的最小音乐行，使用原生 click。
    try:
        result = page.evaluate("""
        () => {
          function visible(el){
            const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
          const nodes = [...document.querySelectorAll('button,[role="button"],div,section,li')].filter(visible);
          const rows = nodes.filter(el => {
            const t = txt(el); const r = el.getBoundingClientRect();
            return t && t.length < 500 && r.width > 220 && r.height > 28 && r.height < 220 &&
              (t.includes('点击添加合适作品风格音乐') || t === '选择音乐');
          }).sort((a,b) => (a.getBoundingClientRect().width*a.getBoundingClientRect().height) -
                            (b.getBoundingClientRect().width*b.getBoundingClientRect().height));
          const row = rows[0];
          if (!row) return {ok:false};
          row.scrollIntoView({block:'center', inline:'nearest'});
          const clickable = [...row.querySelectorAll('button,[role="button"],a,span,div')].filter(visible)
            .find(el => txt(el).includes('选择音乐'));
          (clickable || row).click();
          return {ok:true, text:txt(row).slice(0,120)};
        }
        """)
        if result and result.get("ok") and wait_music_panel_open(page, 3):
            wlog("通过音乐模块 DOM 原生 click 打开面板。")
            return True
    except Exception as exc:
        wlog(f"音乐模块 DOM click 兜底失败：{repr(exc)}")

    # 最后兜底：只点击页面实际“选择音乐”文字的可见矩形，不使用固定屏幕坐标。
    try:
        rects = get_text_rects(page, "选择音乐")
        for rect in rects:
            page.mouse.click(float(rect["cx"]), float(rect["cy"]))
            if wait_music_panel_open(page, 2):
                wlog("通过选择音乐文字矩形打开面板。")
                return True
    except Exception:
        pass

    save_debug(cfg, page, "music_panel_not_opened")
    raise RuntimeError("音乐面板未打开：语义定位、DOM click 和真实鼠标兜底均未生效。")


def _legacy_choose_music(cfg, page):
    """
    V28：强制使用平台音乐面板里的"热门榜"。
    只在右侧音乐抽屉中点击热门榜，并只从右侧音乐列表里悬停歌曲、点击"使用"。
    """
    open_music_panel(cfg, page)
    if verify_music_selected(page):
        return True

    # 只点击右侧抽屉里的"热门榜"
    for attempt in range(1, 5):
        try:
            res = page.evaluate("""
            () => {
              function visible(el){
                const r = el.getBoundingClientRect();
                const s = getComputedStyle(el);
                return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
              }
              function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
              const vw = window.innerWidth;
              const nodes = [...document.querySelectorAll('button,[role=button],span,div,a')].filter(visible);
              const arr = [];
              for (const el of nodes) {
                const t = txt(el);
                if (t !== '热门榜' && !t.includes('热门榜')) continue;
                const r = el.getBoundingClientRect();
                if (r.left < vw * 0.45) continue;  // 只要右侧音乐面板
                if (r.top > 220) continue;         // tab 一般在上方
                let score = 0;
                if (t === '热门榜') score += 100;
                if (r.left > vw * 0.55) score += 20;
                arr.push({el, score, text:t, x:r.left, y:r.top});
              }
              arr.sort((a,b)=>b.score-a.score);
              if (arr.length) {
                (arr[0].el.closest('button') || arr[0].el.closest('[role=button]') || arr[0].el).click();
                return {ok:true, text:arr[0].text};
              }
              return {ok:false};
            }
            """)
            wlog("点击热门榜结果：" + json.dumps(res, ensure_ascii=False)[:300])
            step_wait(cfg, "点击热门榜后等待")
            if res and res.get("ok"):
                break
        except Exception as e:
            wlog(f"第 {attempt} 次点击热门榜失败：{repr(e)}")
        step_wait(cfg, "等待热门榜标签")

    # 从右侧热门榜列表选择音乐
    for attempt in range(1, 9):
        rows = page.evaluate("""
        () => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
          const vw = window.innerWidth, vh = window.innerHeight;
          const out = [];

          // 找右侧音乐面板范围
          const panels = [...document.querySelectorAll('div,section,[role=dialog]')].filter(visible).map(el => {
            const r = el.getBoundingClientRect();
            const t = txt(el);
            return {el, r, t};
          }).filter(x => x.r.left > vw * 0.45 && x.t.includes('选择音乐') && x.t.includes('热门榜'));

          const panel = panels.sort((a,b)=>(b.r.width*b.r.height)-(a.r.width*a.r.height))[0];
          const candidates = panel ? [...panel.el.querySelectorAll('div,li,[role=listitem]')].filter(visible) : [...document.querySelectorAll('div,li,[role=listitem]')].filter(visible);

          for (const el of candidates) {
            const r = el.getBoundingClientRect();
            const t = txt(el);
            if (r.left < vw * 0.45 || r.top < 90 || r.top > vh - 60 || r.width < 250) continue;
            if (r.height < 35 || r.height > 130) continue;
            if (!t || t.length < 5 || t.length > 260) continue;
            if (t.includes('上传') || t.includes('本地') || t.includes('文件') || t.includes('选择音乐')) continue;
            if (t.includes('推荐') && !t.includes('万')) continue;
            // 热门榜行通常有使用量"万使用"和时长
            if (t.includes('万') || /\\d{1,2}:\\d{2}/.test(t)) {
              out.push({
                cx: r.left + r.width * 0.45,
                cy: r.top + r.height / 2,
                text: t.slice(0, 80),
                top: r.top
              });
            }
          }
          return out.slice(0, 15);
        }
        """)

        if not rows:
            wlog("热门榜音乐列表暂未加载，滚动后重试。")
            try:
                page.mouse.wheel(0, 500)
            except Exception:
                pass
            step_wait(cfg, "等待热门榜音乐列表")
            continue

        random.shuffle(rows)
        for row in rows[:6]:
            wlog("悬停热门榜音乐：" + row.get("text", ""))
            page.mouse.move(float(row["cx"]), float(row["cy"]))
            step_wait(cfg, "悬停热门榜音乐后等待使用按钮")

            btn = page.evaluate("""
            () => {
              function visible(el){
                const r = el.getBoundingClientRect();
                const s = getComputedStyle(el);
                return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
              }
              function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
              const vw = window.innerWidth;
              const arr = [];
              for (const el of [...document.querySelectorAll('button,[role=button],span,div')].filter(visible)) {
                const t = txt(el);
                const r = el.getBoundingClientRect();
                if ((t === '使用' || t.startsWith('使用')) && r.left > vw * 0.62 && r.width >= 25 && r.height >= 20) {
                  arr.push({x:r.left+r.width/2, y:r.top+r.height/2, left:r.left});
                }
              }
              arr.sort((a,b)=>b.left-a.left);
              return arr[0] || null;
            }
            """)

            if btn:
                page.mouse.click(float(btn["x"]), float(btn["y"]))
                step_wait(cfg, "点击使用热门榜音乐后等待")
                # 历史兜底也只能确认明确的音乐弹窗，禁止点击整页通用确认按钮。
                confirm_music_dialog_if_present(page)
                step_wait(cfg, "等待热门榜音乐使用生效")
                if verify_music_selected(page):
                    wlog("热门榜音乐已使用成功。")
                    return True

        try:
            page.mouse.wheel(0, 700)
        except Exception:
            pass
        step_wait(cfg, "滚动热门榜音乐列表后等待")

    save_debug(cfg, page, "hot_music_use_failed")
    raise RuntimeError("热门榜音乐面板已打开，但未能点击使用音乐。")


def music_control_belongs_to_panel(locator):
    try:
        return bool(locator.evaluate(r"""
        el => {
          function txt(node){ return ((node.innerText || node.textContent || '') + '').replace(/\s+/g, ' ').trim(); }
          for (let node = el; node && node !== document.body; node = node.parentElement) {
            const text = txt(node);
            const role = node.getAttribute && node.getAttribute('role');
            const data = node.getAttribute && ((node.getAttribute('data-e2e') || '') + (node.getAttribute('data-testid') || ''));
            const r = node.getBoundingClientRect();
            const position = getComputedStyle(node).position;
            const semanticPanel = role === 'dialog' || (node.getAttribute && node.getAttribute('aria-modal') === 'true') || /music/i.test(data);
            const boundedDrawer = r.width < window.innerWidth * 0.72 && r.height < window.innerHeight * 1.35 && r.top >= -50;
            if (text.includes('选择音乐') && (text.includes('使用') || text.includes('热门') || text.includes('推荐')) &&
                (semanticPanel || (boundedDrawer && ['fixed','absolute','sticky'].includes(position))) &&
                !(text.includes('基础信息') && text.includes('发布设置'))) {
              return true;
            }
          }
          return false;
        }
        """))
    except Exception:
        return False


def visible_music_use_buttons(page):
    """返回音乐面板内当前可见的使用按钮，优先语义 role，再兼容原生按钮结构。"""
    candidates = []
    locators = [
        page.get_by_role("button", name=re.compile(r"^使用(?:音乐)?$")),
        page.locator('button,[role="button"]').filter(has_text=re.compile(r"^\s*使用(?:音乐)?\s*$")),
        page.get_by_text("使用", exact=True),
        page.get_by_text("使用音乐", exact=True),
    ]
    seen = set()
    for locator in locators:
        try:
            for index in range(min(locator.count(), 30)):
                item = locator.nth(index)
                if not item.is_visible(timeout=200) or not music_control_belongs_to_panel(item):
                    continue
                box = item.bounding_box(timeout=500)
                key = tuple(round(float(box.get(part, 0)), 1) for part in ("x", "y", "width", "height")) if box else (id(item),)
                if key in seen:
                    continue
                seen.add(key)
                candidates.append(item)
        except Exception:
            pass
    return candidates


def mark_music_song_rows(page):
    """标记音乐面板内可能的歌曲行；不要求“万使用”或固定时长文本。"""
    try:
        return page.evaluate(r"""
        () => {
          function visible(el){
            const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' &&
                   s.visibility !== 'hidden' && Number(s.opacity || 1) > 0;
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').replace(/\s+/g, ' ').trim(); }
          document.querySelectorAll('[data-douyin-music-row-candidate]').forEach(el => el.removeAttribute('data-douyin-music-row-candidate'));

          const panels = [...document.querySelectorAll('[role="dialog"],[aria-modal="true"],[data-e2e*="music" i],[data-testid*="music" i],aside,section,div')]
            .filter(visible).map(el => {
              const text = txt(el); const r = el.getBoundingClientRect();
              const role = el.getAttribute('role') || '';
              const data = (el.getAttribute('data-e2e') || '') + (el.getAttribute('data-testid') || '');
              const position = getComputedStyle(el).position;
              const semanticPanel = role === 'dialog' || el.getAttribute('aria-modal') === 'true' || /music/i.test(data);
              let score = 0;
              if (role === 'dialog') score += 1000;
              if (el.getAttribute('aria-modal') === 'true') score += 800;
              if (/music/i.test(data)) score += 700;
              if (position === 'fixed') score += 300;
              score += Math.min(250, el.querySelectorAll('li,[role="listitem"],img').length * 10);
              return {el,text,r,score,semanticPanel,position};
            }).filter(x => x.text.includes('选择音乐') &&
              (x.text.includes('使用') || x.text.includes('热门') || x.text.includes('推荐') || x.text.includes('搜索音乐')) &&
              x.r.width >= 260 && x.r.height >= 180 && x.text.length < 3000 &&
              !(x.text.includes('基础信息') && x.text.includes('发布设置')) &&
              (x.semanticPanel || (['fixed','absolute','sticky'].includes(x.position) &&
                x.r.width < window.innerWidth * 0.72 && x.r.height < window.innerHeight * 1.35 && x.r.top >= -50)))
            .sort((a,b) => b.score - a.score);
          const panel = panels[0] && panels[0].el;
          if (!panel) return {panel:false, rowCount:0, text:''};

          const excluded = /^(选择音乐|热门榜|热门音乐|推荐|推荐音乐|搜索音乐|原创榜|取消|关闭|确定|确认|完成|使用)$/;
          const raw = [...panel.querySelectorAll('li,[role="listitem"],article,div')].filter(visible).map(el => {
            const r = el.getBoundingClientRect(); const text = txt(el);
            const hasImage = !!el.querySelector('img');
            const hasButton = !!el.querySelector('button,[role="button"]');
            let score = 0;
            if ((el.getAttribute('role') || '') === 'listitem') score += 100;
            if (el.tagName === 'LI' || el.tagName === 'ARTICLE') score += 80;
            if (hasImage) score += 70;
            if (hasButton) score += 30;
            if (r.height >= 44 && r.height <= 120) score += 50;
            score -= Math.abs(r.height - 72) / 3;
            return {el,r,text,hasImage,hasButton,score,area:r.width*r.height};
          }).filter(x => x.r.width >= 150 && x.r.height >= 32 && x.r.height <= 180 &&
              x.text.length >= 2 && x.text.length <= 320 && !excluded.test(x.text) &&
              !x.text.startsWith('选择音乐 热门榜') && !x.text.includes('点击添加合适作品风格音乐'));

          raw.sort((a,b) => b.score - a.score || a.area - b.area);
          const chosen = [];
          for (const item of raw) {
            if (chosen.some(existing => Math.abs(existing.r.top - item.r.top) < 8)) continue;
            chosen.push(item);
            if (chosen.length >= 20) break;
          }
          chosen.forEach((item,index) => item.el.setAttribute('data-douyin-music-row-candidate', String(index)));
          return {panel:true, rowCount:chosen.length, text:txt(panel).slice(0,500)};
        }
        """)
    except Exception as exc:
        return {"panel": False, "rowCount": 0, "text": "", "error": repr(exc)}


def scroll_music_panel_list(page, amount=420):
    """只滚动音乐面板内部列表；DOM 滚动失败后才在面板实际边界内使用鼠标滚轮。"""
    try:
        result = page.evaluate("""
        amount => {
          function visible(el){
            const r = el.getBoundingClientRect(); const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
          const panels = [...document.querySelectorAll('[role="dialog"],[aria-modal="true"],[data-e2e*="music" i],[data-testid*="music" i],aside,section,div')]
            .filter(visible).filter(el => {
              const t = txt(el); const r = el.getBoundingClientRect();
              const role = el.getAttribute('role') || '';
              const data = (el.getAttribute('data-e2e') || '') + (el.getAttribute('data-testid') || '');
              const position = getComputedStyle(el).position;
              const semanticPanel = role === 'dialog' || el.getAttribute('aria-modal') === 'true' || /music/i.test(data);
              return t.includes('选择音乐') && (t.includes('使用') || t.includes('热门') || t.includes('推荐')) &&
                !(t.includes('基础信息') && t.includes('发布设置')) &&
                (semanticPanel || (['fixed','absolute','sticky'].includes(position) && r.width < window.innerWidth * 0.72));
            });
          for (const panel of panels) {
            const scrollables = [panel, ...panel.querySelectorAll('div,ul,ol,section')].filter(visible)
              .filter(el => el.scrollHeight > el.clientHeight + 20)
              .sort((a,b) => (b.scrollHeight-b.clientHeight) - (a.scrollHeight-a.clientHeight));
            if (!scrollables.length) continue;
            const target = scrollables[0];
            const before = target.scrollTop;
            target.scrollBy({top:Number(amount || 420), behavior:'auto'});
            const r = target.getBoundingClientRect();
            return {ok:true, moved:target.scrollTop !== before, x:r.left, y:r.top, width:r.width, height:r.height};
          }
          return {ok:false};
        }
        """, int(amount))
        if result and result.get("ok"):
            return True
    except Exception:
        result = None
    try:
        info = music_panel_info(page)
        if info.get("opened"):
            page.mouse.move(
                float(info.get("x", 0)) + float(info.get("width", 0)) / 2,
                float(info.get("y", 0)) + float(info.get("height", 0)) / 2,
            )
            page.mouse.wheel(0, int(amount))
            return True
    except Exception:
        pass
    return False


def activate_music_category(page):
    """列表为空时才切换推荐/热门分类；分类按钮也限定在音乐面板内。"""
    for role in ("tab", "button"):
        for label in ("推荐音乐", "推荐", "热门榜", "热门音乐"):
            try:
                locator = page.get_by_role(role, name=re.compile(rf"^{re.escape(label)}$"))
                for index in range(min(locator.count(), 6)):
                    item = locator.nth(index)
                    if item.is_visible(timeout=200) and music_control_belongs_to_panel(item):
                        if click_locator_with_fallback(page, item, f"切换音乐分类“{label}”", timeout=1500):
                            page.wait_for_timeout(500)
                            return True
            except Exception:
                pass
    return False


def confirm_music_dialog_if_present(page):
    """只在明确属于音乐的确认弹窗中点击确认，绝不全局点击“确定/完成”。"""
    try:
        dialogs = page.get_by_role("dialog")
        for index in range(min(dialogs.count(), 8)):
            dialog = dialogs.nth(index)
            if not dialog.is_visible(timeout=200):
                continue
            text = dialog.inner_text(timeout=800)
            if "音乐" not in text or not any(word in text for word in ("确认", "使用", "替换")):
                continue
            buttons = dialog.get_by_role("button", name=re.compile(r"^(确认使用|使用|确定|确认)$"))
            for button_index in range(min(buttons.count(), 6)):
                button = buttons.nth(button_index)
                if button.is_visible(timeout=200) and click_locator_with_fallback(page, button, "音乐确认弹窗"):
                    return True
    except Exception:
        pass
    return False


def try_apply_music_use(cfg, page, use_button):
    if not click_locator_with_fallback(page, use_button, "使用音乐"):
        return False
    page.wait_for_timeout(300)
    confirm_music_dialog_if_present(page)
    ok, state = wait_music_selection_applied(page, timeout_seconds=10)
    if ok:
        return True
    save_debug(cfg, page, "music_click_state_not_applied")
    wlog("已点击使用，但音乐状态未在限定时间内生效：" + json.dumps(state, ensure_ascii=False)[:500])
    return False


def choose_music(cfg, page):
    """按网页模块直控音乐：可见使用按钮 → 歌曲行内按钮 → 悬停歌曲行 → 鼠标兜底。"""
    if verify_music_selected(page, log_detail=False):
        return True

    last_row_info = {"panel": False, "rowCount": 0, "text": ""}
    clicked_use = False
    saw_rows = False
    for panel_cycle in range(1, 3):
        open_music_panel(cfg, page)
        if verify_music_selected(page, log_detail=False):
            return True

        # 推荐音乐、新版列表可能默认已有可见“使用”按钮，优先直接点击，不强制切热门榜。
        for use_button in visible_music_use_buttons(page):
            clicked_use = True
            if try_apply_music_use(cfg, page, use_button):
                return True

        # 短轮询等待列表加载；只在确实未加载时做一次面板内滚动。
        deadline = time.time() + 4.0
        while time.time() < deadline:
            last_row_info = mark_music_song_rows(page)
            if int(last_row_info.get("rowCount", 0) or 0) > 0 or visible_music_use_buttons(page):
                break
            page.wait_for_timeout(300)

        for use_button in visible_music_use_buttons(page):
            clicked_use = True
            if try_apply_music_use(cfg, page, use_button):
                return True

        rows = page.locator('[data-douyin-music-row-candidate]')
        try:
            row_count = min(rows.count(), 20)
        except Exception:
            row_count = 0
        saw_rows = saw_rows or row_count > 0

        if row_count == 0:
            try:
                activate_music_category(page)
                last_row_info = mark_music_song_rows(page)
                rows = page.locator('[data-douyin-music-row-candidate]')
                row_count = min(rows.count(), 20)
                saw_rows = saw_rows or row_count > 0
            except Exception:
                pass

        if row_count == 0:
            try:
                scroll_music_panel_list(page, 420)
                page.wait_for_timeout(600)
                wlog("音乐列表尚未加载，已在音乐面板内滚动一次后重新检测。")
                last_row_info = mark_music_song_rows(page)
                rows = page.locator('[data-douyin-music-row-candidate]')
                row_count = min(rows.count(), 20)
                saw_rows = saw_rows or row_count > 0
            except Exception:
                pass

        if row_count == 0:
            save_debug(cfg, page, "music_list_not_loaded")

        for row_index in range(row_count):
            row = rows.nth(row_index)
            try:
                if not row.is_visible(timeout=300):
                    continue
            except Exception:
                continue

            # 先找歌曲行内部已存在的“使用”按钮。
            try:
                row_buttons = row.locator('button,[role="button"]').filter(
                    has_text=re.compile(r"^\s*使用(?:音乐)?\s*$")
                )
                for button_index in range(min(row_buttons.count(), 6)):
                    button = row_buttons.nth(button_index)
                    if button.is_visible(timeout=200):
                        clicked_use = True
                        if try_apply_music_use(cfg, page, button):
                            return True
            except Exception:
                pass

            # 按钮只有悬停才显示时，优先 Playwright hover，再用真实鼠标移动兜底。
            hovered = False
            try:
                row.hover(timeout=1500)
                hovered = True
            except Exception:
                try:
                    box = row.bounding_box(timeout=800)
                    if box:
                        page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                        hovered = True
                except Exception:
                    pass
            if not hovered:
                continue
            page.wait_for_timeout(250)
            for use_button in visible_music_use_buttons(page):
                clicked_use = True
                if try_apply_music_use(cfg, page, use_button):
                    return True

        if verify_music_selected(page, log_detail=False):
            return True
        if not platform_music_panel_opened(page):
            wlog(f"音乐面板在第 {panel_cycle} 次选择中意外关闭且音乐未生效，重新打开后只重试音乐步骤。")
            continue
        break

    if clicked_use:
        save_debug(cfg, page, "music_click_state_not_applied")
        raise RuntimeError("已点击音乐使用按钮，但音乐模块仍未显示歌曲信息或修改入口。")
    if not last_row_info.get("panel"):
        save_debug(cfg, page, "music_panel_not_opened")
        raise RuntimeError("音乐面板未保持打开，无法读取音乐列表。")
    if not saw_rows:
        panel_text = str(last_row_info.get("text") or "")
        debug_name = "music_list_not_loaded" if any(word in panel_text for word in ("加载", "暂无", "重试")) else "music_song_row_not_found"
        save_debug(cfg, page, debug_name)
        raise RuntimeError("音乐面板已打开，但未识别到可用歌曲行；新版列表不要求‘万使用’或固定时长。")
    save_debug(cfg, page, "music_use_button_not_found")
    raise RuntimeError("已识别歌曲行，但行内及悬停后均未找到可用的‘使用’按钮。")


def close_music_panel_if_open(cfg, page):
    if not platform_music_panel_opened(page):
        return
    if not verify_music_selected(page, log_detail=False):
        wlog("音乐尚未确认生效，保留音乐面板供当前音乐步骤继续处理。")
        return
    # 先点音乐面板自己的关闭按钮；只把 Escape 作为最后兜底。
    try:
        buttons = page.get_by_role("button", name=re.compile(r"^(关闭|收起)$"))
        for index in range(min(buttons.count(), 8)):
            button = buttons.nth(index)
            if button.is_visible(timeout=200) and music_control_belongs_to_panel(button):
                if click_locator_with_fallback(page, button, "关闭音乐面板", timeout=1500):
                    page.wait_for_timeout(250)
                    if not platform_music_panel_opened(page):
                        return
    except Exception:
        pass
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(250)
    except Exception:
        pass


def select_music(cfg, page):
    if verify_music_selected(page, log_detail=False):
        wlog("音乐已经生效，跳过打开面板和重复选择。")
        return True
    choose_music(cfg, page)
    if not verify_music_selected(page, log_detail=False):
        validate_music_added(cfg, page)
    return True


def validate_publish_settings_visible(cfg, page, raise_error=True):
    """
    V20：校验"发布设置"区域确实在当前可视区域，而不是仅存在于 DOM 文本里。
    必须看到：
    - 发布设置
    - 发布时间/立即发布/定时发布
    - 底部发布按钮区域
    """
    try:
        info = page.evaluate("""
        () => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }
          function inViewport(r){
            return r.top >= 0 && r.top < window.innerHeight && r.bottom > 0 && r.left < window.innerWidth && r.right > 0;
          }

          let publishSettingVisible = false;
          let publishTimeVisible = false;
          let publishButtonVisible = false;

          const nodes = [...document.querySelectorAll('div,section,span,label,p,button')].filter(visible);
          for (const el of nodes) {
            const t = txt(el);
            if (!t || t.length > 120) continue;
            const r = el.getBoundingClientRect();
            const iv = inViewport(r);

            if (iv && t.includes('发布设置')) publishSettingVisible = true;
            if (iv && (t.includes('发布时间') || t.includes('立即发布') || t.includes('定时发布'))) publishTimeVisible = true;
            if (iv && (t === '发布' || t.includes('发布'))) {
              // 只接受页面底部区域的发布按钮，避免左侧导航"发布"误判
              if (r.top > window.innerHeight * 0.55 && r.left > window.innerWidth * 0.1 && r.left < window.innerWidth * 0.75) {
                publishButtonVisible = true;
              }
            }
          }

          return {
            publishSettingVisible,
            publishTimeVisible,
            publishButtonVisible,
            scrollY: window.scrollY,
            innerHeight: window.innerHeight,
            bodyHeight: document.documentElement.scrollHeight
          };
        }
        """)
    except Exception as e:
        info = {"error": repr(e), "publishSettingVisible": False, "publishTimeVisible": False, "publishButtonVisible": False}

    ok = bool(info.get("publishSettingVisible") and info.get("publishTimeVisible") and info.get("publishButtonVisible"))
    wlog("发布设置可视校验：" + json.dumps(info, ensure_ascii=False)[:500])

    if not ok and raise_error:
        save_debug(cfg, page, "publish_settings_not_visible")
        raise RuntimeError("没有在当前可视区域检测到完整发布设置和底部发布按钮。")
    return ok


def scroll_to_publish_settings(cfg, page):
    """
    V20：强制滚动发布页到最底部。
    抖音页面常见内部容器滚动，不是 window 滚动，所以同时滚动所有可滚动容器。
    只有发布设置和底部发布按钮都进入可视区后，才算成功。
    """
    # 发布设置和底部发布按钮已经在当前视口时，说明页面已经到达需要的
    # 操作位置。此时不要再滚动或等待，避免页面抖动和重复耗时。
    if validate_publish_settings_visible(cfg, page, raise_error=False):
        wlog("当前已经位于发布页底部，跳过下滑和鼠标滚轮步骤。")
        return True

    wlog("发布页尚未到底部，开始下滑到发布设置区域。")

    scroll_js = """
    () => {
      function visible(el){
        const r = el.getBoundingClientRect();
        const s = getComputedStyle(el);
        return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
      }
      function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }

      // 先找发布设置或发布按钮，尽量滚入视野
      const keys = ['发布设置', '发布时间', '立即发布', '定时发布'];
      const nodes = [...document.querySelectorAll('div,section,span,label,p,button')].filter(visible);
      for (const key of keys) {
        const el = nodes.find(n => {
          const t = txt(n);
          return t && t.length < 120 && t.includes(key);
        });
        if (el) {
          el.scrollIntoView({block:'center', inline:'nearest'});
          break;
        }
      }

      // window 到底
      window.scrollTo(0, document.documentElement.scrollHeight);

      // 所有内部滚动容器到底
      const all = [...document.querySelectorAll('*')];
      for (const el of all) {
        try {
          const r = el.getBoundingClientRect();
          const st = getComputedStyle(el);
          const canScroll = el.scrollHeight > el.clientHeight + 60;
          const likelyMain = r.width > 300 && r.height > 300;
          if (canScroll && likelyMain) {
            el.scrollTop = el.scrollHeight;
          }
        } catch(e) {}
      }
      return true;
    }
    """

    for i in range(1, 8):
        try:
            page.evaluate(scroll_js)
        except Exception as e:
            wlog(f"第 {i} 次滚动发布设置失败：{repr(e)}")

        # DOM 滚动通常会立即生效，先短轮询校验；已经到达底部就不再执行
        # 鼠标滚轮和随机长等待。
        try:
            page.wait_for_timeout(250)
        except Exception:
            pass
        if validate_publish_settings_visible(cfg, page, raise_error=False):
            wlog("DOM 下滑后已到达发布页底部，跳过鼠标滚轮步骤。")
            return True

        # DOM 滚动未把目标区域带入视口时，才使用真实鼠标滚轮兜底。
        try:
            page.mouse.wheel(0, 1200)
        except Exception:
            pass

        try:
            page.wait_for_timeout(450)
        except Exception:
            pass

        if validate_publish_settings_visible(cfg, page, raise_error=False):
            wlog("已成功下滑到发布设置区域，底部发布按钮可见。")
            return True

    save_debug(cfg, page, "publish_settings_scroll_failed")
    raise RuntimeError("未能下滑到发布设置区域最底部。")


def get_text_rects(page, word):
    return page.evaluate("""
    (word) => {
      const walker=document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      const out=[];
      while(walker.nextNode()){
        const node=walker.currentNode, text=node.nodeValue||'', idx=text.indexOf(word);
        if(idx<0) continue;
        const range=document.createRange(); range.setStart(node,idx); range.setEnd(node,idx+word.length);
        for(const r of range.getClientRects()){
          if(r.width>0&&r.height>0&&r.top>0&&r.top<window.innerHeight){
            out.push({x:r.left,y:r.top,w:r.width,h:r.height,cx:r.left+r.width/2,cy:r.top+r.height/2});
          }
        }
      }
      out.sort((a,b)=>b.y-a.y);
      return out;
    }
    """, word) or []


def visible_schedule_input_state(page):
    """识别“发布时间”同一模块中已经显示的日期时间输入框。"""
    try:
        return page.evaluate(r"""
        () => {
          function visible(el){
            const r=el.getBoundingClientRect(), s=getComputedStyle(el);
            return r.width>0 && r.height>0 && s.display!=='none' &&
              s.visibility!=='hidden' && Number(s.opacity||1)>0;
          }
          function txt(el){ return ((el.innerText||el.textContent||'')+'').replace(/\s+/g,' ').trim(); }
          const labels=[...document.querySelectorAll('label,span,div,p')]
            .filter(visible).filter(el => txt(el)==='发布时间')
            .map(el => el.getBoundingClientRect());
          const inputs=[...document.querySelectorAll('input')].filter(visible)
            .filter(el => !['radio','checkbox','file','hidden','button','submit']
              .includes(String(el.type||'').toLowerCase()));
          const results=[];
          for(const input of inputs){
            const r=input.getBoundingClientRect();
            if(r.width<90 || r.height<20) continue;
            const value=String(input.value||'').trim();
            const placeholder=String(input.getAttribute('placeholder')||'');
            const dateLike=/\d{4}[-/]\d{1,2}[-/]\d{1,2}\s+\d{1,2}:\d{2}/.test(value);
            const hintLike=/日期|时间|发布/.test(placeholder);
            let nearLabel=false;
            for(const lr of labels){
              const dy=Math.abs((r.top+r.height/2)-(lr.top+lr.height/2));
              if(dy<55 && r.left>lr.left-20 && r.left<lr.right+700){ nearLabel=true; break; }
            }
            let moduleText='';
            for(let parent=input.parentElement, depth=0; parent && depth<6; parent=parent.parentElement, depth++){
              const text=txt(parent);
              if(text.length<500 && text.includes('发布时间') && text.includes('定时发布')){
                moduleText=text;
                break;
              }
            }
            if((nearLabel || moduleText) && (dateLike || hintLike)){
              results.push({value, placeholder, x:r.left, y:r.top, width:r.width, height:r.height});
            }
          }
          return {visible:results.length>0, count:results.length, inputs:results.slice(0,4)};
        }
        """) or {"visible": False, "count": 0, "inputs": []}
    except Exception as exc:
        return {"visible": False, "count": 0, "inputs": [], "error": repr(exc)}


def publish_mode_selected(page, mode):
    word = "定时发布" if mode == "schedule" else "立即发布"
    try:
        radios = page.get_by_role("radio", name=re.compile(rf"^{re.escape(word)}$"))
        for index in range(min(radios.count(), 8)):
            radio = radios.nth(index)
            if radio.is_visible(timeout=200) and radio.is_checked(timeout=300):
                return True
    except Exception:
        pass
    try:
        selected = bool(page.evaluate(r"""
        word => {
          function visible(el){ const r=el.getBoundingClientRect(); const s=getComputedStyle(el); return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden'; }
          function txt(el){ return ((el.innerText||el.textContent||'')+'').replace(/\s+/g,' ').trim(); }
          const labels=[...document.querySelectorAll('label,[role="radio"],button,div,span')].filter(visible)
            .filter(el => { const t=txt(el); return t===word || (t.includes(word) && t.length<80); });
          for(const el of labels){
            const roleRadio=el.closest('[role="radio"]');
            if(roleRadio && (roleRadio.getAttribute('aria-checked')==='true' || roleRadio.getAttribute('data-state')==='checked')) return true;
            const label=el.closest('label');
            let input=label && label.querySelector('input[type="radio"]');
            if(!input && label && label.htmlFor) input=document.getElementById(label.htmlFor);
            if(input && input.checked) return true;
            const er=el.getBoundingClientRect();
            const nearbyChecked=[...document.querySelectorAll('input[type="radio"]:checked')]
              .filter(visible).some(radio => {
                const rr=radio.getBoundingClientRect();
                return Math.abs((rr.top+rr.height/2)-(er.top+er.height/2))<45 &&
                  rr.left>er.left-220 && rr.left<er.right+220;
              });
            if(nearbyChecked) return true;
            const control=el.closest('button,[aria-checked],[data-state]');
            if(control && (control.getAttribute('aria-checked')==='true' || control.getAttribute('data-state')==='checked')) return true;
            const classRoot=label||roleRadio||control||el;
            const cls=([classRoot, ...classRoot.querySelectorAll('*')]
              .map(node => String(node.className||'')).join(' ')).toLowerCase();
            if(/checked|selected|active/.test(cls)) return true;
          }
          return false;
        }
        """, word))
        if selected:
            return True
    except Exception:
        pass

    # 抖音新版定时单选控件可能不暴露 checked/aria-checked，但选中后会在
    # “发布时间”同一行显示可编辑日期时间框。该输入框出现本身就是定时模式
    # 已生效的可验证结果，不能再误报“没有成功点击定时发布”。
    if mode == "schedule":
        return bool(visible_schedule_input_state(page).get("visible"))
    return False


def click_publish_mode(page, mode):
    word = "定时发布" if mode == "schedule" else "立即发布"
    alt = "定时" if mode == "schedule" else "直接发布"
    if publish_mode_selected(page, mode):
        return True

    semantic_locators = [
        page.get_by_role("radio", name=re.compile(rf"^{re.escape(word)}$")),
        page.get_by_text(word, exact=True),
        page.get_by_text(alt, exact=True),
    ]
    for locator in semantic_locators:
        try:
            for index in range(min(locator.count(), 10)):
                item = locator.nth(index)
                if not item.is_visible(timeout=200):
                    continue
                if click_locator_with_fallback(page, item, f"选择{word}", timeout=1800):
                    deadline = time.time() + 3
                    while time.time() < deadline:
                        if publish_mode_selected(page, mode):
                            return True
                        page.wait_for_timeout(200)
        except Exception:
            pass

    # DOM 原生 click 兜底，仍按模块文本/role/label 定位。
    try:
        clicked = page.evaluate("""
        ({word,alt}) => {
          function visible(el){ const r=el.getBoundingClientRect(); const s=getComputedStyle(el); return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden'; }
          function txt(el){ return ((el.innerText||el.textContent||'')+'').trim(); }
          const nodes=[...document.querySelectorAll('label,[role="radio"],button,span,div')].filter(visible)
            .filter(el => { const t=txt(el); return (t===word||t===alt) && t.length<30; });
          const el=nodes[0]; if(!el) return false;
          (el.closest('label')||el.closest('[role="radio"]')||el.closest('button')||el).click();
          return true;
        }
        """, {"word": word, "alt": alt})
        if clicked:
            deadline = time.time() + 2
            while time.time() < deadline:
                if publish_mode_selected(page, mode):
                    return True
                page.wait_for_timeout(200)
    except Exception:
        pass

    # 最后才使用文字实际矩形附近的真实鼠标，不使用固定屏幕坐标。
    for rect in get_text_rects(page, word) or get_text_rects(page, alt):
        for x, y in ((rect["cx"], rect["cy"]), (rect["x"] - 18, rect["cy"])):
            if x < 0:
                continue
            page.mouse.click(float(x), float(y))
            page.wait_for_timeout(250)
            if publish_mode_selected(page, mode):
                return True
    return False


def schedule_input_candidates(page):
    try:
        return page.evaluate("""
        () => {
          function visible(el){ const r=el.getBoundingClientRect(); const s=getComputedStyle(el); return r.width>0&&r.height>0&&s.display!=='none'&&s.visibility!=='hidden'; }
          function txt(el){ return ((el.innerText||el.textContent||'')+'').trim(); }
          const labels=[...document.querySelectorAll('div,span,label,p')].filter(visible);
          const rects=[];
          for(const el of labels){ const t=txt(el); if(t.includes('发布时间')||t.includes('发布设置')){ const r=el.getBoundingClientRect(); rects.push({x:r.left,y:r.top,w:r.width,h:r.height}); } }
          const out=[];
          [...document.querySelectorAll('input')].filter(visible).forEach((el,idx)=>{
            const r=el.getBoundingClientRect(); const type=String(el.type||'').toLowerCase();
            if(['radio','checkbox','file','hidden','button','submit'].includes(type)) return;
            if(r.width<90||r.height<20) return;
            const val=el.value||'', ph=el.getAttribute('placeholder')||'', cls=String(el.className||'');
            let score=0;
            for(const pr of rects){
              const dy=Math.abs((r.top+r.height/2)-(pr.y+pr.h/2));
              if(dy<100 && r.top>pr.y-30) score+=60;
            }
            if(val.includes('202')) score+=35;
            if(ph.includes('日期')||ph.includes('时间')||ph.includes('发布')||cls.includes('picker')) score+=20;
            if(score>0) out.push({idx,x:r.left,y:r.top,w:r.width,h:r.height,value:val,score});
          });
          out.sort((a,b)=>b.score-a.score);
          return out.slice(0,6);
        }
        """) or []
    except Exception:
        return []


def schedule_input_locator(page, candidate):
    best = None
    best_distance = float("inf")
    try:
        inputs = page.locator("input")
        target_x = float(candidate.get("x", 0)) + float(candidate.get("w", 0)) / 2
        target_y = float(candidate.get("y", 0)) + float(candidate.get("h", 0)) / 2
        for index in range(min(inputs.count(), 80)):
            item = inputs.nth(index)
            if not item.is_visible(timeout=150):
                continue
            input_type = str(item.get_attribute("type", timeout=300) or "text").lower()
            if input_type in {"radio", "checkbox", "file", "hidden", "button", "submit"}:
                continue
            box = item.bounding_box(timeout=300)
            if not box or box["width"] < 90 or box["height"] < 20:
                continue
            distance = abs(box["x"] + box["width"] / 2 - target_x) + abs(box["y"] + box["height"] / 2 - target_y)
            if distance < best_distance:
                best = item
                best_distance = distance
    except Exception:
        return None
    return best


def normalize_schedule_input_value(value):
    """把网页时间框常见格式统一为 YYYY-MM-DD HH:MM，便于可靠回读。"""
    raw = str(value or "").strip().replace("/", "-").replace("T", " ")
    match = re.search(
        r"(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})",
        raw,
    )
    if not match:
        return raw
    year, month, day, hour, minute = (int(part) for part in match.groups())
    return f"{year:04d}-{month:02d}-{day:02d} {hour:02d}:{minute:02d}"


def schedule_value_matches(value, slot):
    return normalize_schedule_input_value(value) == slot.strftime("%Y-%m-%d %H:%M")


def type_schedule_time_with_keyboard(page, input_locator, slot):
    """优先用真实键盘写入完整日期时间，并在失焦后读取同一输入框校验。"""
    target = slot.strftime("%Y-%m-%d %H:%M")
    try:
        if input_locator.evaluate("el => Boolean(el.readOnly || el.disabled)", timeout=800):
            wlog("定时时间框为只读控件，直接使用 v3.0.1 DOM 清空与输入逻辑。")
            return False
        input_locator.scroll_into_view_if_needed(timeout=1500)
        input_locator.click(timeout=1800)
        old_value = input_locator.evaluate("el => String(el.value || '')", timeout=1500)
        input_locator.press("Control+A", timeout=1200)
        page.keyboard.press("Backspace")
        page.wait_for_timeout(120)
        cleared_value = input_locator.evaluate("el => String(el.value || '')", timeout=1500)
        wlog(f"清空定时时间框：原值={old_value}；清空后={cleared_value}")
        if cleared_value:
            wlog("键盘清空定时时间框未生效，将按 v3.0.1 的 DOM 原生输入方式兜底。")
            return False
        page.keyboard.type(target, delay=35)
        page.keyboard.press("Tab")
        page.wait_for_timeout(450)
        value = input_locator.evaluate("el => String(el.value || '')", timeout=1500)
        wlog(f"键盘输入定时时间：目标={target}；回读={value}")
        return schedule_value_matches(value, slot)
    except Exception as exc:
        wlog(f"键盘输入定时时间未完成，将使用日期面板兜底：{repr(exc)}")
        return False


def set_schedule_time_with_dom(page, input_locator, slot):
    """按 v3.0.1 的稳定逻辑清空同一 input，再写入完整目标时间并回读。"""
    target = slot.strftime("%Y-%m-%d %H:%M")
    try:
        result = input_locator.evaluate(r"""
        (el, target) => {
          const setter =
            Object.getOwnPropertyDescriptor(el.__proto__, 'value')?.set ||
            Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
          const emitInput = value => {
            try {
              el.dispatchEvent(new InputEvent('input', {
                bubbles:true, inputType:value ? 'insertText' : 'deleteContentBackward', data:value
              }));
            } catch(e) {
              el.dispatchEvent(new Event('input', {bubbles:true}));
            }
          };
          const oldValue=String(el.value||'');
          el.focus();
          if(el._valueTracker) el._valueTracker.setValue(oldValue);
          if(setter) setter.call(el, ''); else el.value='';
          const clearedValue=String(el.value||'');
          emitInput('');
          el.dispatchEvent(new Event('change', {bubbles:true}));
          const valueAfterClearEvents=String(el.value||'');
          if(el._valueTracker) el._valueTracker.setValue(valueAfterClearEvents);
          if(setter) setter.call(el, target); else el.value=target;
          emitInput(target);
          el.dispatchEvent(new Event('change', {bubbles:true}));
          el.dispatchEvent(new KeyboardEvent('keydown', {bubbles:true,key:'Enter',code:'Enter'}));
          el.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true,key:'Enter',code:'Enter'}));
          el.dispatchEvent(new FocusEvent('blur', {bubbles:true}));
          return {oldValue, clearedValue, valueAfterClearEvents, finalValue:String(el.value||''), target};
        }
        """, target, timeout=1800)
        page.wait_for_timeout(450)
        value = input_locator.evaluate("el => String(el.value || '')", timeout=1500)
        wlog("按 v3.0.1 逻辑清空并输入定时时间：" + json.dumps(result, ensure_ascii=False)[:500])
        wlog(f"DOM 输入定时时间回读：目标={target}；回读={value}")
        return schedule_value_matches(value, slot)
    except Exception as exc:
        wlog(f"v3.0.1 DOM 定时时间输入未完成，将使用日期面板兜底：{repr(exc)}")
        return False


def set_schedule_time(page, slot):
    """
    V27：优化定时时间输入。
    日期选择保持原逻辑；时间不再滚动选择小时/分钟。
    根据实测：删除原有时分，再输入目标 HH:MM 即可正常。
    所以流程改为：
    1. 点击发布时间输入框打开面板。
    2. 点击目标日期。
    3. 对同一个时间输入框保留日期部分，删除原有 HH:MM，输入目标 HH:MM。
    4. 点击确认。
    """
    target = slot.strftime("%Y-%m-%d %H:%M")
    target_date = slot.strftime("%Y-%m-%d")
    target_time = slot.strftime("%H:%M")
    day_num = str(slot.day)

    cands = schedule_input_candidates(page)
    if not cands:
        wlog("未找到定时时间输入框。")
        return False

    cand = cands[0]

    # 点击发布时间输入框，打开日期时间面板
    try:
        input_locator = schedule_input_locator(page, cand)
        if input_locator is None or not click_locator_with_fallback(
            page, input_locator, "打开发布时间输入模块", timeout=800
        ):
            raise RuntimeError("发布时间输入模块直接点击失败")
    except Exception as e:
        wlog(f"点击时间输入框失败：{repr(e)}")
        return False

    # 先按 v3.0.1 已在真实页面验证的稳定方式，对同一个受控 input 依次执行
    # 清空、写入完整目标值和回读；该方式能兼容新版日期控件的动画与只读外观。
    if set_schedule_time_with_dom(page, input_locator, slot):
        if verify_schedule_time(page, slot):
            wlog("已按 v3.0.1 的输入逻辑清空、写入并确认定时时间。")
            return True
        wlog("v3.0.1 DOM 输入后的全局回读未匹配，继续使用键盘输入兜底。")

    if type_schedule_time_with_keyboard(page, input_locator, slot):
        if verify_schedule_time(page, slot):
            wlog("已通过键盘输入并确认定时时间。")
            return True
        wlog("键盘输入后的全局回读未匹配，继续使用日期面板兜底。")

    step_wait(reason="直接输入定时时间未生效，等待日期面板兜底")

    # 点击日期，日期这一步当前已经稳定，保留
    try:
        res = page.evaluate("""
        ({day}) => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').trim(); }

          const nodes = [...document.querySelectorAll('td,button,span,div')].filter(visible);
          const arr = [];
          for (const el of nodes) {
            const t = txt(el);
            if (t !== day) continue;
            const r = el.getBoundingClientRect();
            if (r.width > 90 || r.height > 90 || r.top < 160) continue;
            const cls = String(el.className || '').toLowerCase();
            let score = 0;
            if (cls.includes('disabled')) score -= 300;
            if (r.left > 250 && r.left < window.innerWidth * 0.8) score += 20;
            if (cls.includes('selected') || cls.includes('active') || cls.includes('today')) score += 5;
            arr.push({el, score, x:r.left, y:r.top});
          }
          arr.sort((a,b)=>b.score-a.score);
          if (arr.length) {
            arr[0].el.click();
            return {ok:true, count:arr.length};
          }
          return {ok:false, count:0};
        }
        """, {"day": day_num})
        wlog(f"点击日期结果：{res}")
    except Exception as e:
        wlog(f"点击日期失败：{repr(e)}")

    step_wait(reason="点击日期后等待")

    # 核心修复：保留日期，清掉旧时分，再输入目标 HH:MM
    # 不使用全页面 Ctrl+A；只操作最近的发布时间 input。
    try:
        result = page.evaluate("""
        ({x, y, targetDate, targetTime, target}) => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
          }

          const inputs = [...document.querySelectorAll('input')].filter(visible)
            .filter(el => !['radio','checkbox','file','hidden','button','submit'].includes(String(el.type||'').toLowerCase()))
            .filter(el => {
              const r = el.getBoundingClientRect();
              return r.width >= 90 && r.height >= 20;
            });

          let best = null;
          let bestD = 999999;
          for (const el of inputs) {
            const r = el.getBoundingClientRect();
            const cx = r.left + r.width / 2;
            const cy = r.top + r.height / 2;
            const d = Math.abs(cx - x) + Math.abs(cy - y);
            if (d < bestD) {
              bestD = d;
              best = el;
            }
          }

          if (!best) return {ok:false, reason:'input not found', value:''};

          best.scrollIntoView({block:'center'});
          best.focus();
          best.click();

          let current = best.value || '';

          // 先尽量保留已经选好的日期；如果输入框日期不对，直接用目标日期
          let datePart = targetDate;
          const m = current.match(/\\d{4}[-/]\\d{1,2}[-/]\\d{1,2}/);
          if (m) datePart = m[0].replace(/\\//g, '-');

          // 如果当前日期不是目标日期，仍以目标日期为准
          if (datePart !== targetDate) datePart = targetDate;

          const finalValue = datePart + ' ' + targetTime;

          const setter =
            Object.getOwnPropertyDescriptor(best.__proto__, 'value')?.set ||
            Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;

          // 模拟"删除原时分，再输入新时分"：最终只改这个 input 的值，不动页面其他文字
          if (best._valueTracker) best._valueTracker.setValue(current);
          if (setter) setter.call(best, finalValue);
          else best.value = finalValue;

          best.dispatchEvent(new InputEvent('input', {bubbles:true, inputType:'insertText', data:finalValue}));
          best.dispatchEvent(new Event('change', {bubbles:true}));

          // 把光标放在最后，触发一次 Enter
          try {
            best.setSelectionRange(finalValue.length, finalValue.length);
          } catch(e) {}

          best.dispatchEvent(new KeyboardEvent('keydown', {bubbles:true, key:'Enter', code:'Enter'}));
          best.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true, key:'Enter', code:'Enter'}));
          best.dispatchEvent(new FocusEvent('blur', {bubbles:true}));

          return {ok:true, oldValue:current, value:best.value || '', finalValue};
        }
        """, {
            "x": cand["x"] + cand["w"] / 2,
            "y": cand["y"] + cand["h"] / 2,
            "targetDate": target_date,
            "targetTime": target_time,
            "target": target,
        })
        wlog(f"删除原时分并输入目标时间结果：{result}")
    except Exception as e:
        wlog(f"输入目标时分失败：{repr(e)}")
        result = {"ok": False}

    step_wait(reason="输入目标时分后等待")

    # 只在明确属于日期/时间选择器的浮层内点击确认，禁止扫描整页的
    # “确定/确认/完成”，避免误触音乐、封面或其他设置模块。
    try:
        confirm_result = page.evaluate(r"""
        ({targetDate, targetTime}) => {
          function visible(el){
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' &&
                   s.visibility !== 'hidden' && Number(s.opacity || 1) > 0;
          }
          function txt(el){ return ((el.innerText || el.textContent || '') + '').replace(/\s+/g, ' ').trim(); }
          const buttons = [...document.querySelectorAll('button,[role="button"]')]
            .filter(visible)
            .filter(el => /^(确定|确认)$/.test(txt(el)));
          const candidates = [];
          for (const button of buttons) {
            let parent = button.parentElement;
            for (let depth = 0; parent && depth < 8; depth++, parent = parent.parentElement) {
              if (!visible(parent)) continue;
              const text = txt(parent);
              const r = parent.getBoundingClientRect();
              if (text.length > 3000 || r.width < 180 || r.height < 100) continue;
              const looksLikeTimePicker = /发布时间|选择日期|选择时间|时|分/.test(text) ||
                text.includes(targetDate) || text.includes(targetTime) ||
                !!parent.querySelector('input[value*="'+targetDate+'"]');
              const wrongModule = /选择音乐|封面设置|作品描述/.test(text);
              if (looksLikeTimePicker && !wrongModule) {
                candidates.push({button, area:r.width*r.height});
                break;
              }
            }
          }
          candidates.sort((a,b) => a.area-b.area);
          if (!candidates.length) return {clicked:false, reason:'schedule confirmation not found'};
          candidates[0].button.click();
          return {clicked:true};
        }
        """, {"targetDate": target_date, "targetTime": target_time})
        if confirm_result and confirm_result.get("clicked"):
            step_wait(reason="确认定时时间后等待")
    except Exception as exc:
        wlog(f"定时时间选择器确认按钮处理失败，将直接读取输入值校验：{repr(exc)}")

    return verify_schedule_time(page, slot)


def verify_schedule_time(page, slot):
    target = slot.strftime("%Y-%m-%d %H:%M")
    vals = [str(x.get("value","")).strip() for x in schedule_input_candidates(page)]
    wlog(f"定时校验：目标={target}；输入框值={vals}")
    return any(schedule_value_matches(value, slot) for value in vals)


def set_schedule(cfg, page, slot):
    scroll_to_publish_settings(cfg, page)
    if not click_publish_mode(page, "schedule"):
        save_debug(cfg, page, "schedule_mode_failed")
        raise RuntimeError("没有成功点击定时发布。")
    deadline = time.time() + 4
    mode_state = visible_schedule_input_state(page)
    while not mode_state.get("visible") and time.time() < deadline:
        page.wait_for_timeout(150)
        mode_state = visible_schedule_input_state(page)
    if not mode_state.get("visible"):
        save_debug(cfg, page, "schedule_input_not_visible")
        raise RuntimeError("定时发布已点击，但发布时间输入框没有出现。")
    wlog("定时发布已选中，准备输入目标时间：" + json.dumps(mode_state, ensure_ascii=False)[:500])
    for _ in range(2):
        if set_schedule_time(page, slot):
            wlog("定时时间设置成功。")
            return
        step_wait(cfg, "定时时间未成功，重试")
    save_debug(cfg, page, "schedule_time_failed")
    raise RuntimeError("定时时间没有成功设置。")


def set_publish_now(cfg, page):
    """
    不定时发布逻辑：
    这里只是切换到"立即发布/不定时发布"的单选状态。
    最终提交仍然点击底部红色"发布"按钮，不存在单独的"立即发布"提交按钮。
    """
    scroll_to_publish_settings(cfg, page)
    if not click_publish_mode(page, "now"):
        save_debug(cfg, page, "publish_now_failed")
        raise RuntimeError("没有成功切换到立即发布/不定时发布状态。")
    step_wait(cfg, "切换立即发布状态后等待")



def publish_page_still_visible(page):
    """
    v1.1 修复：判断是否仍停留在发布编辑页。
    点击发布后不再点"确认/确定"，只观察页面是否跳转。
    """
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""
    if "content/post/image" in url or "creator-micro/content/upload" in url:
        return True

    try:
        body = page.locator("body").inner_text(timeout=1500)
    except Exception:
        body = ""

    publish_markers = ["基础信息", "作品描述", "发布设置", "选择音乐", "封面设置"]
    manage_markers = ["作品管理", "全部作品", "已发布", "审核中", "未通过"]

    if any(x in body for x in manage_markers) and not any(x in body for x in publish_markers[:2]):
        return False

    return any(x in body for x in publish_markers)


def manage_page_visible(page):
    """
    v1.1 修复：点击发布后，成功状态以进入作品管理页/作品列表为准。
    """
    try:
        url = (page.url or "").lower()
    except Exception:
        url = ""

    if "content/manage" in url or "creator-micro/content/manage" in url:
        return True

    try:
        body = page.locator("body").inner_text(timeout=1500)
    except Exception:
        body = ""

    if "作品管理" in body and ("全部作品" in body or "已发布" in body or "审核中" in body or "未通过" in body):
        return True

    # 定时作品发布后列表里常见"定时发布中"
    if "定时发布中" in body and ("作品管理" in body or "修改定时" in body):
        return True

    return False


def find_bottom_publish_button(page):
    """
    找底部真正的红色"发布"按钮。
    不找"立即发布"按钮，因为不定时发布也是点击底部"发布"。
    """
    try:
        buttons = page.get_by_role("button").filter(has_text="发布")
        arr = []
        for i in range(min(buttons.count(), 30)):
            try:
                b = buttons.nth(i)
                if not b.is_visible(timeout=300):
                    continue
                box = b.bounding_box(timeout=500)
                if not box:
                    continue
                txt = (b.inner_text(timeout=500) or "").strip()
                normalized = re.sub(r"\s+", "", txt)
                if normalized in {"立即发布", "定时发布", "重新发布", "发布设置"}:
                    continue
                if normalized not in {"发布", "发布作品", "确认发布"}:
                    continue
                button_type = str(b.get_attribute("type", timeout=300) or "").lower()
                data_hint = " ".join(filter(None, (
                    b.get_attribute("data-e2e", timeout=300),
                    b.get_attribute("data-testid", timeout=300),
                    b.get_attribute("aria-label", timeout=300),
                ))).lower()
                score = float(box.get("y", 0))
                if button_type == "submit":
                    score += 2000
                if "publish" in data_hint or "submit" in data_hint:
                    score += 1500
                if normalized == "发布":
                    score += 300
                arr.append((score, b, txt))
            except Exception:
                pass
        if not arr:
            return None
        arr.sort(key=lambda x: x[0])
        return arr[-1][1]
    except Exception:
        return None


def click_bottom_publish_button(cfg, page):
    scroll_to_publish_settings(cfg, page)
    btn = find_bottom_publish_button(page)
    if btn is None:
        save_debug(cfg, page, "publish_button_not_found")
        raise RuntimeError("没有找到底部发布按钮。")
    try:
        if btn.is_disabled(timeout=500) or btn.get_attribute("aria-disabled", timeout=500) == "true":
            raise RuntimeError("底部发布按钮当前不可用。")
    except PlaywrightTimeoutError:
        pass
    if not click_locator_with_fallback(page, btn, "底部发布按钮", timeout=8000):
        save_debug(cfg, page, "publish_button_click_failed")
        raise RuntimeError("底部发布按钮语义点击、DOM click 和真实鼠标兜底均失败。")
    step_wait(cfg, "点击发布后等待")
    return True


def wait_publish_result_once(cfg, page, wait_seconds=35):
    """
    点击发布后等待一次结果。
    返回：
    success：已经确认进入作品管理页
    still_publish：仍停留在发布页，需要再次点击发布
    error：检测到失败提示
    unknown：未确定状态
    """
    end = time.time() + wait_seconds
    last_state_log = 0
    success_notice_seen = False

    while time.time() < end:
        if daily_publish_limit_visible(page):
            save_debug(cfg, page, "daily_publish_limit")
            wlog("发布结果判断：检测到‘今天投稿次数已达到上限，请明天再试’。")
            return "daily_limit"

        if manage_page_visible(page):
            wlog("发布结果判断：已进入作品管理页面，发布成功。")
            return "success"

        if has_any(page, ["发布成功", "定时发布成功", "提交成功", "作品已进入定时发布", "发布任务已提交"], 600):
            if not success_notice_seen:
                wlog("发布结果判断：检测到成功提示，继续等待并确认进入作品管理页。")
                success_notice_seen = True

        if has_any(page, ["发布失败", "错误", "请完善", "不能为空", "违规", "过于频繁", "稍后再试"], 600):
            save_debug(cfg, page, "publish_error")
            wlog("发布结果判断：检测到失败或错误提示。")
            return "error"

        if publish_page_still_visible(page):
            if time.time() - last_state_log > 8:
                wlog("发布结果判断：仍停留在发布页面，继续等待跳转。")
                last_state_log = time.time()
        else:
            if time.time() - last_state_log > 8:
                try:
                    wlog(f"发布结果判断：页面跳转中，当前URL={page.url}")
                except Exception:
                    wlog("发布结果判断：页面跳转中。")
                last_state_log = time.time()

        time.sleep(2)

    if daily_publish_limit_visible(page):
        save_debug(cfg, page, "daily_publish_limit")
        return "daily_limit"

    if manage_page_visible(page):
        return "success"

    if publish_page_still_visible(page):
        return "still_publish"

    if success_notice_seen:
        wlog("发布结果判断：虽出现成功提示，但未确认进入作品管理页，因此不计为成功。")
    return "unknown"


def submit(cfg, page):
    """
    单次提交并判断结果。完整作品重试只由 worker 外层控制，
    保证前端“失败重试次数”不会被内外两层重复放大。
    """
    wlog("点击一次底部发布按钮，并等待进入作品管理页。")
    click_bottom_publish_button(cfg, page)
    result = wait_publish_result_once(cfg, page, wait_seconds=35)
    if result == "success" or manage_page_visible(page):
        return True
    if result == "daily_limit":
        raise DailyPublishLimitError("抱歉，今天投稿次数已达到上限，请明天再试")
    if result == "still_publish":
        save_debug(cfg, page, "publish_still_on_publish_page")
        wlog("发布结果判断：仍停留在发布页面，本次完整作品尝试失败。")
    elif result == "error":
        wlog("发布结果判断：页面返回失败提示，本次完整作品尝试失败。")
    else:
        save_debug(cfg, page, "publish_result_unknown")
        wlog("发布结果判断：未能确认是否进入作品管理页，本次完整作品尝试失败。")
    return False



def keep_single_browser_tab(context, page=None):
    """
    V29：强制单标签模式。
    优先保留当前传入标签；如果没有，则保留最后一个创作者中心标签。
    关闭其它标签，避免越跑越多。
    """
    try:
        pages = list(context.pages)
        if not pages:
            return context.new_page()

        chosen = page if page in pages else None
        if chosen is None:
            douyin_pages = []
            for p in pages:
                try:
                    if "creator.douyin.com" in (p.url or ""):
                        douyin_pages.append(p)
                except Exception:
                    pass
            chosen = douyin_pages[-1] if douyin_pages else pages[-1]

        for p in list(pages):
            if p is chosen:
                continue
            try:
                p.close()
            except Exception:
                pass

        # v1.1：不主动拉起/恢复最小化浏览器
        safe_bring_to_front(ACTIVE_CFG or {}, chosen)
        wlog("已整理浏览器标签：仅保留 1 个发布标签。")
        return chosen
    except Exception as e:
        wlog(f"整理浏览器标签失败：{repr(e)}")
        try:
            return page or context.pages[-1]
        except Exception:
            return page




def wait_publish_interval(cfg, done, remaining_slots=None):
    """
    V1.1：每条发布成功后，进入下一条前等待一个可浮动发布间隔。
    这个等待和"每一步等待秒"分开，专门用于两条作品之间的间隔。
    """
    if remaining_slots is not None and remaining_slots <= 0:
        return
    try:
        a = float(cfg.get("publish_interval_min_seconds", 30) or 0)
        b = float(cfg.get("publish_interval_max_seconds", 90) or 0)
    except Exception:
        a, b = 30, 90

    if a < 0:
        a = 0
    if b < a:
        b = a

    if b <= 0:
        return

    d = random.uniform(a, b) if b > a else a
    wlog(f"发布间隔等待 {d:.1f} 秒后进入下一条。")
    countdown_sleep(cfg, d, "发布间隔：准备下一条", "publish_interval")


def account_worker(config_path):
    global ACTIVE_CFG
    cfg = json.loads(Path(config_path).read_text(encoding="utf-8"))
    accounts = browser_accounts_from_config(cfg, enabled_only=True)
    if not accounts:
        raise RuntimeError("没有启用任何浏览器账号，请在发布中心至少启用一个浏览器。")
    validate_browser_queue_ports(accounts)

    # 独立子进程只处理一个账号；先把该账号绑定的内容池映射成有效配置。
    cfg = config_for_browser_account(cfg, accounts[0])
    cfg["browser_accounts"] = [accounts[0]]
    accounts = [accounts[0]]
    ACTIVE_CFG = cfg
    records = read_publish_records(cfg)
    slots = build_slots(cfg) if cfg.get("use_schedule", True) else []
    state = read_state(cfg)
    primary_account = accounts[0]
    primary_quota = max(1, int(primary_account.get("posts_per_run", 1) or 1))
    # 轮换发布模式下，父进程仍保留账号完整配额，但每次只允许当前账号成功发布指定条数。
    # 默认值 0 表示沿用原模式：一次处理完当前账号的全部配额。
    try:
        queue_pass_limit = max(0, int(cfg.get("_queue_pass_limit", 0) or 0))
    except Exception:
        queue_pass_limit = 0
    queue_run_id = str(cfg.get("_queue_run_id") or f"direct_{uuid.uuid4().hex}")
    resume_progress = account_run_resume_progress(state, queue_run_id, primary_quota)
    current_schedule_signature = schedule_signature(cfg)
    if state.get("schedule_signature") != current_schedule_signature:
        old_si = int(state.get("slot_index", 0) or 0)
        state["slot_index"] = 0
        state["schedule_signature"] = current_schedule_signature
        write_state(cfg, state)
        wlog(
            f"检测到排期配置已变化，定时时间点从 {old_si + 1} 调整为 1；"
            f"本轮已完成 {resume_progress}/{primary_quota} 条继续保留。"
        )

    state.update({
        "queue_run_id": queue_run_id,
        "run_progress": resume_progress,
        "run_quota": primary_quota,
        "run_active": resume_progress < primary_quota,
        "last_account_outcome": "",
        "last_account_message": "",
    })
    write_state(cfg, state)

    ci = 0 if cfg.get("delete_copy_after_success", True) else int(state.get("copy_index", 0) or 0)
    si = int(state.get("slot_index", 0) or 0)
    retry_times = normalize_retry_count(cfg.get("retry_times", 3), 3)
    overall_retry_times = normalize_retry_count(cfg.get("overall_retry_times", 1), 1)
    total_done = 0
    pass_done = 0
    stop_all = False

    wlog(
        f"本轮启用 {len(accounts)} 个浏览器账号；单步骤失败后最多重试 {retry_times} 次，"
        f"单步骤仍失败时总体流程最多重试 {overall_retry_times} 次。"
    )
    if queue_pass_limit > 0:
        wlog(f"账号轮换子任务：本次最多成功发布 {queue_pass_limit} 条，完成后立即切换下一个账号。")
    wlog(f"读取发布记录 {len(records)} 条，当前记录序号：{ci + 1}")
    if cfg.get("use_schedule", True):
        _slot, current_cycle, current_offset = repeating_schedule_slot(slots, si)
        wlog(
            f"独立排期 {len(slots)} 个，当前为第 {current_cycle} 轮、"
            f"轮内时间点 {current_offset + 1}/{len(slots)}。"
        )
    else:
        wlog("此队列项使用立即发布；发布条数只取账号队列配置。")

    for account_index, account in enumerate(accounts, start=1):
        if stop_all:
            break

        account_cfg = config_for_browser_account(cfg, account)
        ACTIVE_CFG = account_cfg
        account_name = account.get("name") or f"浏览器{account_index}"
        account_quota = max(1, int(account.get("posts_per_run", 1) or 1))
        account_done = min(account_quota, resume_progress)
        launch_info = None
        browser = None
        playwright_manager = None
        daily_limit_hit = False
        daily_limit_message = ""

        wlog("=" * 60)
        wlog(f"切换到 {account_name}（{account_index}/{len(accounts)}），本账号计划发布 {account_quota} 条。")
        if account_done:
            wlog(
                f"已恢复上一轮进度：{account_name} 已完成 {account_done}/{account_quota} 条，"
                f"本次继续剩余 {max(0, account_quota - account_done)} 条。"
            )
        if account_cfg.get("use_schedule", True):
            _slot, current_cycle, current_offset = repeating_schedule_slot(slots, si)
            wlog(
                f"{account_name} 使用独立队列进度；本轮条数只取账号队列配置，"
                f"当前从第 {current_cycle} 轮、轮内时间点 {current_offset + 1}/{len(slots)} 开始。"
            )
        else:
            wlog(f"{account_name} 使用独立队列进度；本轮条数只取账号队列配置，当前为立即发布。")
        if account_done >= account_quota:
            completed_state = read_state(account_cfg)
            completed_state.update({
                "queue_run_id": queue_run_id,
                "run_progress": account_quota,
                "run_quota": account_quota,
                "run_active": False,
            })
            write_state(account_cfg, completed_state)
            emit_browser_status(account, "done", "上一轮中本账号配额已经完成", account_done, account_quota)
            wlog(f"{account_name} 在上一轮中已经完成，无需重新启动浏览器。")
            continue

        emit_browser_status(account, "launching", "正在启动并连接浏览器", account_done, account_quota)

        if account_cfg.get("use_schedule", True) and account_quota > len(slots):
            wlog(
                f"{account_name} 本轮计划 {account_quota} 条、生成 {len(slots)} 个时间点；"
                "时间点用完后将从本轮第一个时间点继续循环。"
            )

        try:
            playwright_manager = sync_playwright()
            playwright = playwright_manager.start()
            launch_info = open_browser_with_cdp(account_cfg, force_new=False)
            browser = playwright.chromium.connect_over_cdp(
                f"http://127.0.0.1:{int(account_cfg['cdp_port'])}"
            )
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            page = context.pages[-1] if context.pages else context.new_page()
            page = keep_single_browser_tab(context, page)
            safe_bring_to_front(account_cfg, page)

            emit_browser_status(
                account, "checking", "浏览器已连接；发布当前作品时检查登录状态和发布页",
                account_done, account_quota,
            )

            while account_done < account_quota and (queue_pass_limit <= 0 or pass_done < queue_pass_limit):
                records = read_publish_records(account_cfg)
                if account_cfg.get("delete_copy_after_success", True):
                    records = filter_used_publish_records(records, read_state(account_cfg))
                    ci = 0
                if not records:
                    alert_auto_pause("提醒：Excel 中已没有可发布文案，全部任务结束。")
                    append_log(account_cfg, {
                        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "image": "", "copy_index": 0,
                        "schedule_time": "", "copy_preview": "",
                        "status": f"paused_no_copies[{account_name}]"
                    })
                    stop_all = True
                    break
                if ci >= len(records):
                    alert_auto_pause("提醒：文案序号超过当前文案数量，请重置发布进度。")
                    stop_all = True
                    break
                publish_record = records[ci]
                city = publish_record.get("city", "")
                title = publish_record.get("title", "")
                copy_text = publish_record["copy"]
                try:
                    img = choose_image_for_record(account_cfg, publish_record)
                except Exception as exc:
                    alert_auto_pause(f"自动暂停提醒：{account_name} 的图片匹配失败：{exc}")
                    append_log(account_cfg, {
                        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "image": "", "copy_index": ci + 1, "schedule_time": "",
                        "copy_preview": copy_text[:80].replace("\n", " "),
                        "status": f"paused_image_match[{account_name}]: {repr(exc)}"
                    })
                    stop_all = True
                    break
                schedule_cycle = 1
                schedule_offset = 0
                if account_cfg.get("use_schedule", True):
                    slot, schedule_cycle, schedule_offset = repeating_schedule_slot(slots, si)
                    if schedule_offset == 0 and si > 0:
                        wlog(f"{account_name} 定时时间点已循环到第 {schedule_cycle} 轮。")
                else:
                    slot = datetime.now()
                ok = False
                last_error = None
                emit_browser_status(
                    account, "publishing",
                    f"正在发布第 {account_done + 1}/{account_quota} 条",
                    account_done, account_quota
                )

                def publish_current_record(workflow_attempt, workflow_total):
                    """执行当前作品；每个 run_publish_step 只重试自己尚未完成的步骤。"""
                    nonlocal page
                    wlog("-" * 60)
                    wlog(
                        f"{account_name} 当前作品总体流程 {workflow_attempt}/{workflow_total}："
                        f"城市={city or '旧格式根目录'}；图片={img.name}；文案序号={ci + 1}；"
                        f"定时={slot.strftime('%Y-%m-%d %H:%M') if account_cfg.get('use_schedule', True) else '立即发布'}"
                    )

                    page = run_publish_step(
                        account_cfg, page, "整理浏览器标签",
                        lambda: keep_single_browser_tab(context, page),
                    )
                    run_publish_step(
                        account_cfg, page, "进入图文发布页",
                        lambda: enter_publish_page(account_cfg, page),
                    )
                    run_publish_step(
                        account_cfg, page, "上传图片",
                        lambda: upload_image(account_cfg, page, img),
                    )

                    def wait_for_editor():
                        step_wait(account_cfg, "已选择图片，等待进入图文编辑页")
                        return wait_after_upload_to_editor(account_cfg, page)

                    run_publish_step(account_cfg, page, "等待进入图文编辑页", wait_for_editor)

                    def input_and_validate_title():
                        fill_title(account_cfg, page, title)
                        return validate_title_filled(account_cfg, page, title)

                    def input_and_validate_copy():
                        fill_copy(account_cfg, page, copy_text)
                        return validate_copy_filled(account_cfg, page, copy_text)

                    run_publish_step(account_cfg, page, "输入并校验标题", input_and_validate_title)
                    run_publish_step(account_cfg, page, "输入并校验正文话题", input_and_validate_copy)

                    if account_cfg.get("music_required", True):
                        def select_and_validate_music():
                            # 若页面仍显示未选择音乐，本步骤会重新拉起音乐面板并选择；
                            # 只有本步骤连续失败后才进入总体重试并重新上传当前作品。
                            select_music(account_cfg, page)
                            return validate_music_added(account_cfg, page)

                        run_publish_step(account_cfg, page, "选择并校验音乐", select_and_validate_music)

                    def locate_publish_settings():
                        close_music_panel_if_open(account_cfg, page)
                        scroll_to_publish_settings(account_cfg, page)
                        return validate_publish_settings_visible(account_cfg, page)

                    run_publish_step(account_cfg, page, "定位发布设置", locate_publish_settings)

                    if account_cfg.get("use_schedule", True):
                        def set_and_verify_schedule():
                            set_schedule(account_cfg, page, slot)
                            if not verify_schedule_time(page, slot):
                                raise RuntimeError("发布前定时时间校验失败。")
                            return True

                        run_publish_step(account_cfg, page, "设置并校验定时发布", set_and_verify_schedule)
                    else:
                        # 未勾选定时发布时，抖音图文页默认就是立即发布。
                        # 不再点击或校验“立即发布”单选框，直接进入发布前校验并
                        # 点击底部真正的“发布”按钮，避免新版自定义单选控件误判。
                        wlog("当前账号未启用定时发布：跳过发布模式设置，直接准备点击底部发布按钮。")

                    def final_validation():
                        wait_uploaded_with_config(account_cfg, page)
                        validate_uploaded_image(account_cfg, page)
                        validate_title_filled(account_cfg, page, title)
                        validate_copy_filled(account_cfg, page, copy_text)
                        if account_cfg.get("music_required", True):
                            validate_music_added(account_cfg, page)
                        if account_cfg.get("use_schedule", True) and not verify_schedule_time(page, slot):
                            raise RuntimeError("点击发布前定时时间回读不一致，禁止提交。")
                        return True

                    run_publish_step(account_cfg, page, "发布前最终校验", final_validation)

                    def submit_and_confirm():
                        # 上一次点击已成功但页面响应较慢时，不重复点击发布。
                        if manage_page_visible(page):
                            return True
                        if not submit(account_cfg, page):
                            raise RuntimeError("提交发布后未确认进入作品管理页。")
                        return True

                    return run_publish_step(account_cfg, page, "提交并确认作品管理页", submit_and_confirm)

                def prepare_overall_retry(failed_attempt, failed_step_error):
                    nonlocal page
                    save_debug(
                        account_cfg, page,
                        f"{account.get('id', 'browser')}_workflow_{failed_attempt}_{failed_step_error.step_name}_failed",
                    )
                    try:
                        page = keep_single_browser_tab(context, page)
                        page.goto(account_cfg["creator_url"], wait_until="domcontentloaded", timeout=30000)
                    except Exception as exc:
                        wlog(f"总体重试前返回创作者中心失败，将由‘进入图文发布页’步骤继续处理：{repr(exc)}")

                try:
                    ok = bool(run_publish_workflow(account_cfg, publish_current_record, prepare_overall_retry))
                except DailyPublishLimitError as exc:
                    daily_limit_hit = True
                    daily_limit_message = str(exc) or "今天投稿次数已达到上限，请明天再试"
                    last_error = exc
                    wlog(f"{account_name} 检测到今日投稿次数上限：{daily_limit_message}；不进入普通重试。")
                    save_debug(account_cfg, page, f"{account.get('id', 'browser')}_daily_limit")
                except PublishWorkflowFailedError as exc:
                    last_error = exc
                    wlog(f"{account_name} 当前作品两级重试均已耗尽：{exc}")

                if daily_limit_hit:
                    wait_seconds = normalize_daily_limit_switch_wait_seconds(
                        account_cfg.get("daily_limit_switch_wait_seconds", 10)
                    )
                    state_now = read_state(account_cfg)
                    state_now.update({
                        "copy_index": ci,
                        "slot_index": si,
                        "queue_run_id": queue_run_id,
                        "run_progress": account_done,
                        "run_quota": account_quota,
                        "run_active": True,
                        "last_account_outcome": "daily_limit",
                        "last_account_message": daily_limit_message,
                        "daily_limit_detected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                    write_state(account_cfg, state_now)
                    append_log(account_cfg, {
                        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "image": str(img), "copy_index": ci + 1,
                        "schedule_time": slot.strftime("%Y-%m-%d %H:%M") if account_cfg.get("use_schedule", True) else "立即发布",
                        "copy_preview": copy_text[:80].replace("\n", " "),
                        "status": f"daily_publish_limit[{account_name}]: {daily_limit_message}"
                    })
                    emit_browser_status(
                        account, "daily_limit",
                        f"检测到今日投稿上限，{wait_seconds} 秒后切换下一个账号",
                        account_done, account_quota,
                    )
                    wlog(
                        f"{account_name} 今日投稿次数已达到上限；当前内容未消费、已发布数不增加。"
                        f"等待 {wait_seconds} 秒后自动切换下一个账号。"
                    )
                    if wait_seconds > 0:
                        countdown_sleep(
                            account_cfg, wait_seconds,
                            f"{account_name} 今日投稿上限：准备切换下一个账号",
                            "daily_limit_switch",
                        )
                    emit_browser_status(
                        account, "daily_limit", "今日投稿上限，已跳过本账号",
                        account_done, account_quota,
                    )
                    break

                if not ok:
                    failure_reason = concise_publish_failure(last_error)
                    failure_state = read_state(account_cfg)
                    failure_state.update({
                        "copy_index": ci,
                        "slot_index": si,
                        "queue_run_id": queue_run_id,
                        "run_progress": account_done,
                        "run_quota": account_quota,
                        "run_active": True,
                        "last_account_outcome": "error",
                        "last_account_message": failure_reason,
                        "last_account_failed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                    write_state(account_cfg, failure_state)
                    append_log(account_cfg, {
                        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "image": str(img), "copy_index": ci + 1,
                        "schedule_time": slot.strftime("%Y-%m-%d %H:%M") if account_cfg.get("use_schedule", True) else "立即发布",
                        "copy_preview": copy_text[:80].replace("\n", " "),
                        "status": f"error_after_step_and_workflow_retries[{account_name}]: {repr(last_error)}"
                    })
                    emit_browser_status(
                        account, "error",
                        f"{failure_reason}；进度 {account_done}/{account_quota} 已保留",
                        account_done, account_quota,
                    )
                    wlog(
                        f"{account_name} 已用完单步骤重试 {retry_times} 次和总体流程重试 "
                        f"{overall_retry_times} 次；保留当前内容并切换下一个浏览器。"
                    )
                    alert_auto_pause(
                        f"自动暂停提醒：{account_name} 的失败步骤已重试 {retry_times} 次，"
                        f"总体流程又重试 {overall_retry_times} 次后仍未完成，"
                        "当前内容未消费；请查看固定运行日志和调试截图。"
                    )
                    break

                # 只有确认进入作品管理页后，才消费图片、文案和排期。
                mark_publish_record_used_in_state(account_cfg, publish_record)
                if account_cfg.get("delete_image_after_success", True):
                    try:
                        img.unlink()
                        wlog(f"已删除已发布图片：{img.name}")
                    except Exception as exc:
                        wlog(f"删除图片失败：{exc}")
                if account_cfg.get("delete_copy_after_success", True):
                    try:
                        delete_copy_from_excel(account_cfg, publish_record)
                    except Exception as exc:
                        wlog(f"删除已用文案失败：{repr(exc)}")

                append_log(account_cfg, {
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "image": str(img), "copy_index": ci + 1,
                    "schedule_time": slot.strftime("%Y-%m-%d %H:%M") if account_cfg.get("use_schedule", True) else "立即发布",
                    "copy_preview": copy_text[:80].replace("\n", " "),
                    "status": f"success[{account_name}]"
                })

                ci = 0 if account_cfg.get("delete_copy_after_success", True) else ci + 1
                si += 1
                account_done += 1
                pass_done += 1
                total_done += 1
                state_now = read_state(account_cfg)
                state_now.update({
                    "copy_index": ci,
                    "slot_index": si,
                    "used_copy_keys": state_now.get("used_copy_keys", []),
                    "schedule_signature": schedule_signature(account_cfg),
                    "queue_run_id": queue_run_id,
                    "run_progress": account_done,
                    "run_quota": account_quota,
                    "run_active": account_done < account_quota,
                    "last_account_outcome": "success",
                    "last_account_message": "",
                })
                write_state(account_cfg, state_now)
                emit_browser_status(
                    account, "publishing" if account_done < account_quota else "done",
                    f"本账号已成功 {account_done}/{account_quota} 条",
                    account_done, account_quota
                )

                # 普通顺序模式继续在账号内部等待下一条；轮换模式由父进程切换账号后统一控制发布间隔。
                pass_limit_reached = queue_pass_limit > 0 and pass_done >= queue_pass_limit
                more_needed = (account_done < account_quota or account_index < len(accounts)) and not pass_limit_reached
                if more_needed:
                    wait_publish_interval(account_cfg, total_done)

            if daily_limit_hit:
                wlog(f"{account_name} 因今日投稿次数上限结束本次账号任务，等待父调度器切换。")
            elif account_done >= account_quota:
                wlog(f"{account_name} 已完成本账号配额：{account_done}/{account_quota}。")
                emit_browser_status(account, "done", "本账号配额完成", account_done, account_quota)
            elif queue_pass_limit > 0 and pass_done > 0:
                wlog(f"{account_name} 本次轮换已完成 {pass_done} 条，累计 {account_done}/{account_quota}，切换下一个账号。")
                emit_browser_status(account, "waiting", f"本轮已发 {pass_done} 条，准备切换下一个账号", account_done, account_quota)

        except AccountLoggedOutError as exc:
            wlog(f"{account_name} 掉号或需要验证：{exc} 已跳过并切换下一个浏览器。")
            logout_reason = f"账号掉线/需要验证：{exc}"
            logout_state = read_state(account_cfg)
            logout_state.update({
                "copy_index": ci,
                "slot_index": si,
                "queue_run_id": queue_run_id,
                "run_progress": account_done,
                "run_quota": account_quota,
                "run_active": True,
                "last_account_outcome": "logged_out",
                "last_account_message": logout_reason,
                "last_account_failed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            write_state(account_cfg, logout_state)
            emit_browser_status(
                account, "logged_out",
                f"{logout_reason}；进度 {account_done}/{account_quota} 已保留",
                account_done, account_quota,
            )
            try:
                append_log(account_cfg, {
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "image": "", "copy_index": ci + 1, "schedule_time": "",
                    "copy_preview": "", "status": f"browser_logged_out[{account_name}]"
                })
            except Exception:
                pass
        except Exception as exc:
            startup_reason = f"启动或页面检查失败：{exc}"
            startup_state = read_state(account_cfg)
            startup_state.update({
                "copy_index": ci,
                "slot_index": si,
                "queue_run_id": queue_run_id,
                "run_progress": account_done,
                "run_quota": account_quota,
                "run_active": True,
                "last_account_outcome": "error",
                "last_account_message": startup_reason,
                "last_account_failed_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            write_state(account_cfg, startup_state)
            wlog(f"{account_name} {startup_reason} 已切换下一个浏览器，进度已保留。")
            emit_browser_status(
                account, "error",
                f"{startup_reason}；进度 {account_done}/{account_quota} 已保留",
                account_done, account_quota,
            )
        finally:
            if launch_info and launch_info.get("launched") and account.get("close_after_finish", True):
                close_browser_launched_by_app(browser, launch_info, account_name)
            elif launch_info and launch_info.get("reused"):
                wlog(f"{account_name} 使用的是已存在调试连接，未主动关闭，避免误关用户浏览器。")
            if playwright_manager is not None:
                try:
                    playwright_manager.stop()
                except Exception:
                    pass

    ACTIVE_CFG = cfg
    wlog(f"全部浏览器处理完成。本轮成功发布 {total_done} 条。")


def normalize_rotate_batch_size(value):
    """轮换模式每个账号连续成功发布多少条后切换；最少 1 条。"""
    try:
        return max(1, int(value or 1))
    except Exception:
        return 1


def normalize_daily_limit_switch_wait_seconds(value):
    """检测到今日投稿上限后，等待多少秒再切换下一个账号；允许 0 秒立即切换。"""
    try:
        return max(0, min(86400, int(float(value or 0))))
    except Exception:
        return 10


def rotation_batch_target(quota, before_progress, batch_size):
    """计算当前账号这一轮实际最多发布多少条。"""
    quota = max(1, _nonnegative_state_int(quota, 1))
    before_progress = min(quota, _nonnegative_state_int(before_progress, 0))
    remaining = max(0, quota - before_progress)
    if remaining <= 0:
        return 0
    return min(remaining, normalize_rotate_batch_size(batch_size))


def run_isolated_browser_sequence(config_path, child_switch, publish_mode=False):
    """
    每个浏览器使用独立 Python/Playwright 进程，单号驱动崩溃不会影响后续账号。

    普通模式：按队列顺序，一次完成当前账号的全部配额后再切换。
    轮换模式：每个账号每轮最多成功发布 rotate_batch_size 条，然后立即切换下一个账号；
    循环执行，直到所有账号达到各自 posts_per_run 配额。
    """
    cfg = json.loads(Path(config_path).read_text(encoding="utf-8"))
    accounts = browser_accounts_from_config(cfg, enabled_only=True)
    if not accounts:
        raise RuntimeError("没有启用任何浏览器账号。")
    validate_browser_queue_ports(accounts)
    total_done = 0
    queue_run_id = ""
    completed_account_ids = set()
    daily_limit_account_ids = set()
    failed_account_ids = set()
    rotate_each_post = bool(publish_mode and parse_bool(cfg.get("rotate_accounts_each_post"), False))
    # 未勾选轮换发布时不读取、不使用连续发布条数。
    rotate_batch_size = normalize_rotate_batch_size(cfg.get("rotate_batch_size", 1)) if rotate_each_post else 0
    if publish_mode:
        queue_run_id, completed_account_ids, resumed = prepare_queue_run(cfg, accounts)
        if resumed:
            wlog(
                f"继续上一轮队列任务：已有 {len(completed_account_ids)}/{len(accounts)} 个账号处理完成，"
                "其余账号从各自保存进度继续。"
            )
        else:
            wlog(f"开始新的队列轮次：共 {len(accounts)} 个已勾选账号。")
        if rotate_each_post:
            wlog(f"已启用账号轮换发布：每个账号连续成功发布 {rotate_batch_size} 条后切换下一个，直至各账号设定条数全部完成。")
    temp_root = Path(tempfile.mkdtemp(prefix="douyin_browser_sequence_"))
    live_queue_fingerprint = json.dumps(accounts, ensure_ascii=False, sort_keys=True)

    def reload_live_queue():
        """在账号子进程边界重新读取主配置，使暂停期间的队列修改可在继续后生效。"""
        nonlocal cfg, accounts, live_queue_fingerprint, rotate_batch_size
        try:
            latest_cfg = json.loads(Path(config_path).read_text(encoding="utf-8"))
            latest_accounts = browser_accounts_from_config(latest_cfg, enabled_only=True)
            if latest_accounts:
                validate_browser_queue_ports(latest_accounts)
            latest_fingerprint = json.dumps(latest_accounts, ensure_ascii=False, sort_keys=True)
            changed = latest_fingerprint != live_queue_fingerprint
            old_batch_size = rotate_batch_size
            cfg = latest_cfg
            accounts = latest_accounts
            live_queue_fingerprint = latest_fingerprint
            rotate_batch_size = (
                normalize_rotate_batch_size(cfg.get("rotate_batch_size", rotate_batch_size or 1))
                if rotate_each_post else 0
            )
            if changed or rotate_batch_size != old_batch_size:
                detail = ""
                if rotate_batch_size != old_batch_size:
                    detail = f" 连续发布条数已从 {old_batch_size} 更新为 {rotate_batch_size}。"
                wlog("已读取暂停期间保存的最新浏览器队列；后续未启动账号将按新顺序/新配置执行。" + detail)
            return changed
        except Exception as exc:
            wlog(f"读取运行中的最新浏览器队列失败，暂时继续使用原队列：{exc}")
            return False

    def current_account_progress(base_cfg, account, quota):
        resolved_cfg = config_for_browser_account(base_cfg, account)
        state = read_state(resolved_cfg)
        if not publish_mode:
            return 0, resolved_cfg, state
        return account_run_resume_progress(state, queue_run_id, quota), resolved_cfg, state

    def run_one_account(index, account, round_no=1):
        nonlocal total_done
        account_id = str(account.get("id") or "")
        quota = max(1, int(account.get("posts_per_run", 1) or 1))
        before_progress, resolved_child_cfg, _before_state = current_account_progress(cfg, account, quota)
        if publish_mode and before_progress >= quota:
            completed_account_ids.add(account_id)
            emit_browser_status(account, "done", "本账号配额已经完成", quota, quota)
            return 0, True

        child_cfg = copy.deepcopy(cfg)
        child_account = copy.deepcopy(account)
        # 保留完整配额，以便账号 worker 的 run_progress 能跨轮次累计。
        child_account["posts_per_run"] = quota
        child_cfg["browser_accounts"] = [child_account]
        if publish_mode:
            child_cfg["_queue_run_id"] = queue_run_id
            if rotate_each_post:
                # 每个账号连续成功发布指定条数后，子进程必须结束并把控制权交还父进程，
                # 由父进程明确切换到队列中的下一个账号。
                # 本轮上限取“用户设置的连续条数”和“当前账号剩余配额”的较小值。
                # 这样即使总配额不足一个完整批次，也会在完成剩余条数后立刻交棒/完成。
                batch_target = rotation_batch_target(quota, before_progress, rotate_batch_size)
                child_cfg["_queue_pass_limit"] = batch_target
                child_cfg["_rotation_batch_target"] = batch_target
                # 轮换模式下，每个账号第一次轮到时正常打开可见浏览器；
                # 本轮结束后不主动关闭，下一轮直接复用同一 CDP 端口，因此不会反复弹窗。
                # 浏览器窗口由用户在整轮任务结束后自行关闭。
                child_account["close_after_finish"] = False
                child_cfg["browser_accounts"] = [child_account]
                child_cfg["_headless_browser"] = False
                child_cfg["reuse_existing_cdp"] = True
                child_cfg["close_chrome_before_start"] = False
                child_cfg["no_raise_browser"] = True
        child_path = temp_root / f"account_{index}_round_{round_no}.json"
        child_path.write_text(json.dumps(child_cfg, ensure_ascii=False, indent=2), encoding="utf-8")

        if rotate_each_post:
            wlog(
                f"轮换第 {round_no} 轮 · 账号 {index}/{len(accounts)}：{account['name']} "
                f"（当前 {before_progress}/{quota}，本次最多 {rotation_batch_target(quota, before_progress, rotate_batch_size)} 条）"
            )
        else:
            wlog(f"独立进程 {index}/{len(accounts)}：{account['name']}")

        command = application_command(child_switch, child_path)
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        child = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            cwd=str(APP_DIR),
            env=env,
            **hidden_subprocess_kwargs(),
        )
        for line in child.stdout:
            print(line, end="", flush=True)
        code = child.wait()
        try:
            child_path.unlink()
        except Exception:
            pass

        # 如果用户在暂停期间编辑了浏览器队列，当前账号安全结束后立即读取最新配置。
        # 当前正在执行的这个账号仍使用启动它时的参数；尚未启动的账号使用最新参数。
        reload_live_queue()

        if code != 0:
            failed_account_ids.add(account_id)
            emit_browser_status(account, "error", f"独立控制进程退出码 {code}，继续下一个浏览器", before_progress, quota)
            wlog(f"{account['name']} 独立控制进程异常退出（{code}），继续下一个浏览器。")

        if not publish_mode:
            return 0, code == 0

        after_progress, _resolved_cfg, _after_state = current_account_progress(cfg, account, quota)
        gained = max(0, after_progress - before_progress)
        total_done += gained
        if str(_after_state.get("last_account_outcome") or "") == "daily_limit":
            daily_limit_account_ids.add(account_id)
            emit_browser_status(
                account, "daily_limit", "今日投稿上限，已在本轮跳过",
                after_progress, quota,
            )
            wlog(f"{account['name']} 已标记为‘今日投稿上限’，本轮后续不再重新进入该账号。")
            update_queue_run_state(cfg, queue_run_id, accounts, completed_account_ids, active=True)
            return gained, False

        last_outcome = str(_after_state.get("last_account_outcome") or "")
        last_message = str(_after_state.get("last_account_message") or "").strip()
        if last_outcome in {"error", "logged_out"}:
            failed_account_ids.add(account_id)
            status_code = "logged_out" if last_outcome == "logged_out" else "error"
            emit_browser_status(
                account, status_code,
                (last_message or "本账号未完成") + f"；进度 {after_progress}/{quota} 已保留",
                after_progress, quota,
            )
            update_queue_run_state(cfg, queue_run_id, accounts, completed_account_ids, active=True)
            return gained, False

        completed = after_progress >= quota

        if completed:
            completed_account_ids.add(account_id)
            emit_browser_status(account, "done", "本账号配额完成", after_progress, quota)
        update_queue_run_state(cfg, queue_run_id, accounts, completed_account_ids, active=True)
        return gained, completed

    try:
        if rotate_each_post and publish_mode:
            round_no = 0
            while True:
                reload_live_queue()
                pending_now = [
                    account for account in accounts
                    if str(account.get("id") or "") not in completed_account_ids
                    and str(account.get("id") or "") not in daily_limit_account_ids
                    and str(account.get("id") or "") not in failed_account_ids
                ]
                if not pending_now:
                    break

                round_no += 1
                round_progress = 0
                visited_this_round = set()
                while True:
                    reload_live_queue()
                    candidates = [
                        account for account in accounts
                        if str(account.get("id") or "") not in completed_account_ids
                        and str(account.get("id") or "") not in daily_limit_account_ids
                        and str(account.get("id") or "") not in failed_account_ids
                        and str(account.get("id") or "") not in visited_this_round
                    ]
                    if not candidates:
                        break
                    account = candidates[0]
                    account_id = str(account.get("id") or "")
                    visited_this_round.add(account_id)
                    index = next(
                        (i for i, item in enumerate(accounts, start=1)
                         if str(item.get("id") or "") == account_id),
                        1,
                    )
                    gained, account_completed = run_one_account(index, account, round_no)
                    round_progress += gained
                    reload_live_queue()
                    still_pending = any(
                        str(item.get("id") or "") not in completed_account_ids
                        and str(item.get("id") or "") not in daily_limit_account_ids
                        and str(item.get("id") or "") not in failed_account_ids
                        for item in accounts
                    )
                    if gained > 0 and still_pending:
                        # 轮换模式达到“每账号连续发布 N 条”后必须立即交棒。
                        # 条与条之间的发布间隔已经在账号 worker 内部处理；账号边界不再追加等待。
                        wlog(
                            f"{account.get('name') or '当前账号'} 本轮新增 {gained} 条，"
                            "已释放当前账号控制权，立即切换下一个账号。"
                        )
                        if not account_completed:
                            quota_now = max(1, int(account.get("posts_per_run", 1) or 1))
                            progress_now = current_account_progress(cfg, account, quota_now)[0]
                            emit_browser_status(
                                account, "waiting",
                                f"本轮已发 {gained} 条，正在切换下一个账号",
                                progress_now, quota_now,
                            )

                reload_live_queue()
                remaining = [
                    account for account in accounts
                    if str(account.get("id") or "") not in completed_account_ids
                    and str(account.get("id") or "") not in daily_limit_account_ids
                    and str(account.get("id") or "") not in failed_account_ids
                ]
                if not remaining:
                    break
                if round_progress <= 0:
                    wlog(
                        f"轮换第 {round_no} 轮没有任何账号新增成功发布；"
                        "为避免无限循环，已停止本轮并保留各账号进度。"
                    )
                    break
                wlog(
                    f"轮换第 {round_no} 轮结束：本轮新增 {round_progress} 条，"
                    f"还有 {len(remaining)}/{len(accounts)} 个账号未达到设定条数。"
                )
        else:
            attempted_ids = set()
            while True:
                reload_live_queue()
                candidates = [
                    account for account in accounts
                    if str(account.get("id") or "") not in completed_account_ids
                    and str(account.get("id") or "") not in daily_limit_account_ids
                    and str(account.get("id") or "") not in failed_account_ids
                    and str(account.get("id") or "") not in attempted_ids
                ]
                if not candidates:
                    break
                account = candidates[0]
                account_id = str(account.get("id") or "")
                attempted_ids.add(account_id)
                quota = max(1, int(account.get("posts_per_run", 1) or 1))
                index = next(
                    (i for i, item in enumerate(accounts, start=1)
                     if str(item.get("id") or "") == account_id),
                    1,
                )
                if publish_mode and account_id in completed_account_ids:
                    wlog(f"独立进程 {index}/{len(accounts)}：{account['name']} 在上一轮已经处理，跳过重复发布。")
                    emit_browser_status(account, "done", "上一轮已经处理，继续下一个账号", quota, quota)
                    continue
                run_one_account(index, account, 1)

    finally:
        try:
            temp_root.rmdir()
        except Exception:
            pass

    if publish_mode:
        reload_live_queue()
        remaining_ids = {
            str(account.get("id") or "") for account in accounts
            if str(account.get("id") or "") not in completed_account_ids
        }
        update_queue_run_state(
            cfg, queue_run_id, accounts, completed_account_ids, active=bool(remaining_ids)
        )
        runnable_remaining_ids = remaining_ids - daily_limit_account_ids
        if runnable_remaining_ids:
            raise RuntimeError(
                f"本轮仍有 {len(runnable_remaining_ids)} 个账号未处理完成，进度已保留；下次启动将继续。"
            )
        if daily_limit_account_ids:
            wlog(
                f"本轮有 {len(daily_limit_account_ids)} 个账号检测到今日投稿次数上限，"
                "已保留未消费内容和发布进度；下次启动可再次尝试。"
            )
        if remaining_ids:
            wlog(f"本次队列处理结束，新增确认成功 {total_done} 条；未完成账号进度继续保留。")
        elif rotate_each_post:
            wlog(f"账号轮换发布结束，本次新增确认成功 {total_done} 条；所有账号配额均已完成。")
        else:
            wlog(f"多浏览器独立进程编排结束，本次新增确认成功 {total_done} 条，本轮队列已完成。")
    else:
        wlog("多浏览器独立登录检查结束。")


def worker(config_path):
    run_isolated_browser_sequence(config_path, "--account-worker", publish_mode=True)


def probe_browser_accounts(config_path):
    run_isolated_browser_sequence(config_path, "--probe-account", publish_mode=False)


def probe_account_in_process(config_path):
    """只检查全部浏览器能否进入图文发布页，不上传、不填写、不发布。"""
    global ACTIVE_CFG
    cfg = json.loads(Path(config_path).read_text(encoding="utf-8"))
    accounts = browser_accounts_from_config(cfg, enabled_only=True)
    if not accounts:
        raise RuntimeError("没有启用任何浏览器账号。")
    validate_browser_queue_ports(accounts)
    ready_count = 0
    for index, account in enumerate(accounts, start=1):
        account_cfg = config_for_browser_account(cfg, account)
        ACTIVE_CFG = account_cfg
        launch_info = None
        browser = None
        playwright_manager = None
        quota = int(account.get("posts_per_run", 1) or 1)
        emit_browser_status(account, "launching", "登录检查：正在启动浏览器", 0, quota)
        wlog(f"登录检查 {index}/{len(accounts)}：{account['name']}")
        try:
            playwright_manager = sync_playwright()
            playwright = playwright_manager.start()
            launch_info = open_browser_with_cdp(account_cfg, force_new=False)
            browser = playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{int(account_cfg['cdp_port'])}")
            context = browser.contexts[0] if browser.contexts else browser.new_context()
            page = keep_single_browser_tab(context, context.pages[-1] if context.pages else context.new_page())
            emit_browser_status(account, "checking", "登录检查：正在进入发布页", 0, quota)
            enter_publish_page(account_cfg, page)
            ready_count += 1
            emit_browser_status(account, "ready", "登录有效，可以进入图文发布页", 0, quota)
            wlog(f"{account['name']}：登录有效。")
        except AccountLoggedOutError as exc:
            emit_browser_status(account, "logged_out", "检测到登录或安全验证页", 0, quota)
            wlog(f"{account['name']}：掉号或需要验证，{exc}")
        except Exception as exc:
            emit_browser_status(account, "error", f"页面检查失败：{exc}", 0, quota)
            wlog(f"{account['name']}：页面检查失败但未确认掉号，{repr(exc)}")
        finally:
            if launch_info and launch_info.get("launched") and account.get("close_after_finish", True):
                close_browser_launched_by_app(browser, launch_info, account.get("name", "浏览器"))
            if playwright_manager is not None:
                try:
                    playwright_manager.stop()
                except Exception:
                    pass
    ACTIVE_CFG = cfg
    wlog(f"全部账号登录检查完成：{ready_count}/{len(accounts)} 个可进入发布页。")


WEEKDAY_LABELS = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]


def normalized_weekly_days(value):
    if not isinstance(value, (list, tuple, set)):
        return []
    result = []
    for item in value:
        try:
            day = int(item)
        except Exception:
            continue
        if 0 <= day <= 6 and day not in result:
            result.append(day)
    return sorted(result)


def parse_clock_text(value):
    match = re.fullmatch(r"\s*(\d{1,2}):(\d{1,2})\s*", str(value or ""))
    if not match:
        raise ValueError("自动启动时间必须是 HH:MM 格式。")
    hour, minute = int(match.group(1)), int(match.group(2))
    if not 0 <= hour <= 23 or not 0 <= minute <= 59:
        raise ValueError("自动启动时间必须在 00:00 到 23:59 之间。")
    return hour, minute


def weekly_trigger_key(cfg, now=None):
    now = now or datetime.now()
    if not parse_bool(cfg.get("weekly_start_enabled"), False):
        return ""
    days = normalized_weekly_days(cfg.get("weekly_start_days", []))
    if now.weekday() not in days:
        return ""
    hour, minute = parse_clock_text(cfg.get("weekly_start_time", "09:00"))
    if now.hour != hour or now.minute != minute:
        return ""
    return f"{now:%Y-%m-%d} {hour:02d}:{minute:02d}"


def next_weekly_trigger(cfg, now=None):
    now = now or datetime.now()
    if not parse_bool(cfg.get("weekly_start_enabled"), False):
        return None
    days = normalized_weekly_days(cfg.get("weekly_start_days", []))
    if not days:
        return None
    hour, minute = parse_clock_text(cfg.get("weekly_start_time", "09:00"))
    for offset in range(0, 8):
        day = (now + timedelta(days=offset)).date()
        candidate = datetime.combine(day, datetime.min.time()).replace(hour=hour, minute=minute)
        if candidate.weekday() in days and candidate > now:
            return candidate
    return None


class BrowserAccountDialog:
    def __init__(self, owner, account, title="浏览器账号"):
        self.result = None
        self.window = tk.Toplevel(owner)
        self.window.title(title)
        self.window.geometry("820x720")
        self.window.resizable(False, False)
        self.window.transient(owner)
        self.window.grab_set()
        self.vars = {
            "name": tk.StringVar(value=str(account.get("name", ""))),
            "browser_path": tk.StringVar(value=str(account.get("browser_path", ""))),
            "browser_user_data_dir": tk.StringVar(value=str(account.get("browser_user_data_dir", ""))),
            "browser_profile_directory": tk.StringVar(value=str(account.get("browser_profile_directory", ""))),
            "cdp_port": tk.StringVar(value=str(account.get("cdp_port", 9222))),
            "posts_per_run": tk.StringVar(value=str(account.get("posts_per_run", 1))),
            "excel_path": tk.StringVar(value=str(account.get("excel_path", ""))),
            "sheet_name": tk.StringVar(value=str(account.get("sheet_name", ""))),
            "publish_start_date": tk.StringVar(value=str(account.get("publish_start_date", ""))),
            "publish_end_date": tk.StringVar(value=str(account.get("publish_end_date", ""))),
            "start_hour": tk.StringVar(value=str(account.get("start_hour", 8))),
            "end_hour": tk.StringVar(value=str(account.get("end_hour", 18))),
            "per_hour_count": tk.StringVar(value=str(account.get("per_hour_count", 3))),
            "custom_minutes": tk.StringVar(value=str(account.get("custom_minutes", ""))),
        }
        self.enabled_var = tk.BooleanVar(value=parse_bool(account.get("enabled"), True))
        self.close_var = tk.BooleanVar(value=parse_bool(account.get("close_after_finish"), True))
        self.use_schedule_var = tk.BooleanVar(value=parse_bool(account.get("use_schedule"), True))
        self.original_id = str(account.get("id", ""))
        self._build()
        self.window.protocol("WM_DELETE_WINDOW", self.window.destroy)
        self.window.wait_window()

    def _build(self):
        outer = ttk.Frame(self.window, padding=20)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="浏览器队列项配置", font=("Microsoft YaHei UI", 17, "bold")).pack(anchor="w")
        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True, pady=(14, 10))
        account_tab = ttk.Frame(notebook, padding=16)
        schedule_tab = ttk.Frame(notebook, padding=16)
        notebook.add(account_tab, text="浏览器与内容")
        notebook.add(schedule_tab, text="独立发布时间")

        form = ttk.Frame(account_tab)
        form.pack(fill="x")
        rows = [
            ("账号名称", "name", None),
            ("快捷方式 / EXE", "browser_path", "file"),
            ("用户数据目录（可空）", "browser_user_data_dir", "folder"),
            ("Profile 目录（可空）", "browser_profile_directory", None),
            ("CDP 调试端口", "cdp_port", None),
            ("本轮发布条数", "posts_per_run", None),
        ]
        for row, (label, key, browse) in enumerate(rows):
            ttk.Label(form, text=label, width=20, anchor="e").grid(row=row, column=0, padx=(0, 10), pady=7, sticky="e")
            ttk.Entry(form, textvariable=self.vars[key], width=52).grid(row=row, column=1, pady=7, sticky="we")
            if browse:
                ttk.Button(form, text="选择", command=lambda k=key, b=browse: self._browse(k, b)).grid(row=row, column=2, padx=(8, 0), pady=7)
        form.columnconfigure(1, weight=1)
        checks = ttk.Frame(account_tab)
        checks.pack(fill="x", pady=(12, 8))
        ttk.Checkbutton(checks, text="启用此浏览器账号", variable=self.enabled_var).pack(side="left", padx=(150, 18))
        ttk.Checkbutton(checks, text="完成后关闭本软件启动的浏览器", variable=self.close_var).pack(side="left")
        content = ttk.LabelFrame(account_tab, text="账号发布内容", padding=12)
        content.pack(fill="x", pady=(8, 8))
        content_rows = [
            ("文案 Excel", "excel_path", "excel"),
            ("工作表名", "sheet_name", None),
        ]
        for row, (label, key, browse) in enumerate(content_rows):
            ttk.Label(content, text=label, width=16, anchor="e").grid(
                row=row, column=0, padx=(0, 10), pady=5, sticky="e"
            )
            ttk.Entry(content, textvariable=self.vars[key], width=52).grid(
                row=row, column=1, pady=5, sticky="we"
            )
            if browse:
                ttk.Button(
                    content,
                    text="选择",
                    command=lambda k=key, b=browse: self._browse(k, b),
                ).grid(row=row, column=2, padx=(8, 0), pady=5)
        content.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            schedule_tab,
            text="使用抖音定时发布（未勾选时立即发布）",
            variable=self.use_schedule_var,
        ).pack(anchor="w", pady=(0, 12))
        schedule_form = ttk.Frame(schedule_tab)
        schedule_form.pack(fill="x")
        schedule_rows = [
            ("开始日期", "publish_start_date"),
            ("结束日期", "publish_end_date"),
            ("开始小时", "start_hour"),
            ("结束小时 0-24", "end_hour"),
            ("每小时条数", "per_hour_count"),
            ("分钟点", "custom_minutes"),
        ]
        for row, (label, key) in enumerate(schedule_rows):
            ttk.Label(schedule_form, text=label, width=20, anchor="e").grid(
                row=row, column=0, padx=(0, 10), pady=8, sticky="e"
            )
            ttk.Entry(schedule_form, textvariable=self.vars[key], width=52).grid(
                row=row, column=1, pady=8, sticky="we"
            )
        schedule_form.columnconfigure(1, weight=1)

        actions = ttk.Frame(outer)
        actions.pack(fill="x")
        ttk.Button(actions, text="取消", command=self.window.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(actions, text="保存账号", style="Accent.TButton", command=self._save).pack(side="right")

    def _browse(self, key, kind):
        if kind == "folder":
            value = filedialog.askdirectory(parent=self.window)
        elif kind == "excel":
            value = filedialog.askopenfilename(
                parent=self.window,
                title="选择此账号的文案 Excel",
                filetypes=[("Excel 文案表", "*.xlsx *.xls *.xlsm")],
            )
        else:
            value = filedialog.askopenfilename(
                parent=self.window,
                title="选择浏览器快捷方式或程序",
                filetypes=[("浏览器", "*.lnk *.exe"), ("全部文件", "*.*")],
            )
        if value:
            self.vars[key].set(value)

    def _save(self):
        try:
            name = self.vars["name"].get().strip()
            browser_path = self.vars["browser_path"].get().strip()
            port = int(self.vars["cdp_port"].get().strip())
            posts = int(self.vars["posts_per_run"].get().strip())
            if not name:
                raise ValueError("请填写账号名称。")
            if not browser_path:
                raise ValueError("请选择浏览器快捷方式或 EXE。")
            if not 1 <= port <= 65535:
                raise ValueError("CDP 调试端口必须在 1—65535 之间。")
            if posts < 1:
                raise ValueError("每个浏览器本轮至少发布 1 条。")
            excel_path = self.vars["excel_path"].get().strip()
            sheet_name = self.vars["sheet_name"].get().strip()
            if not excel_path or not sheet_name:
                raise ValueError("文案 Excel 和工作表名都必须填写。")
            validate_copy_file_path(excel_path)
            raw = {
                "id": self.original_id,
                "name": name,
                "enabled": self.enabled_var.get(),
                "browser_path": browser_path,
                "browser_user_data_dir": self.vars["browser_user_data_dir"].get().strip(),
                "browser_profile_directory": self.vars["browser_profile_directory"].get().strip(),
                "cdp_port": port,
                "posts_per_run": posts,
                "close_after_finish": self.close_var.get(),
                "independent_content": True,
                "excel_path": excel_path,
                "sheet_name": sheet_name,
                "use_schedule": self.use_schedule_var.get(),
                "publish_start_date": self.vars["publish_start_date"].get().strip(),
                "publish_end_date": self.vars["publish_end_date"].get().strip(),
                "start_hour": int(self.vars["start_hour"].get().strip() or 0),
                "end_hour": int(self.vars["end_hour"].get().strip() or 0),
                "per_hour_count": int(self.vars["per_hour_count"].get().strip() or 0),
                "custom_minutes": self.vars["custom_minutes"].get().strip(),
            }
            self.result = normalize_browser_account(raw, 0, {})
            if self.original_id:
                self.result["id"] = self.original_id
            if self.result.get("use_schedule", True):
                build_slots(config_for_browser_account({}, self.result))
            self.window.destroy()
        except Exception as exc:
            messagebox.showwarning("账号配置无效", str(exc), parent=self.window)


class App:
    def __init__(self, root):
        self.root = root
        # 授权恢复完成前不展示任何发布配置，避免主界面先闪现再弹登录框。
        self.root.withdraw()
        self.root.title(APP_NAME + "｜智能发布中心")
        self.root.geometry("1240x880")
        self.root.minsize(1080, 760)
        try:
            if LOGO_ICO.exists():
                self.root.iconbitmap(str(LOGO_ICO))
        except Exception:
            pass
        self.root.configure(bg="#EEF2F7")
        self.cfg = load_config()
        self.debug_cleanup_result = (
            {"launch_count": 0, "cleaned": False, "removed": 0, "error": ""}
            if "--self-test" in sys.argv
            else maintain_debug_screenshots(self.cfg)
        )
        cleanup_legacy_install_after_migration()
        self.browser_accounts = browser_accounts_from_config(self.cfg)
        self.vars = {}
        self.bool_vars = {}
        self.weekly_day_vars = []
        self.browser_statuses = {}
        self.browser_success_counts = {}
        self.restore_browser_runtime_state()
        self.proc = None
        self.proc_mode = ""
        self.q = queue.Queue()
        self.is_paused = False
        # 暂停期间允许编辑队列。保存暂停时实际创建的 flag 路径，避免删除/改动队列后
        # “继续”只清理新队列 flag，导致旧 worker 仍卡在暂停状态。
        self.active_pause_flags = []
        self.status_var = tk.StringVar(value="状态：空闲")
        self.auth_status_var = tk.StringVar(value="授权状态：正在检查")
        self.auth_button_text_var = tk.StringVar(value="登录授权")
        self.countdown_var = tk.StringVar(value="倒计时：--:--")
        self.countdown_reason_var = tk.StringVar(value="等待任务开始")
        self.next_schedule_var = tk.StringVar(value="下次自动启动：未启用")
        self.auth_client = PlatformAuthClient(
            AUTH_API_BASE_URL,
            AUTH_PRODUCT_CODE,
            APP_VERSION,
            AUTH_TOKEN_PATH,
            AUTH_RELEASE_CHANNEL,
        )
        self.login_preference_store = DpapiTokenStore(AUTH_LOGIN_PREFERENCES_PATH)
        self.auth_heartbeat_inflight = False
        self.auth_force_logout_prompt_shown = False
        self.telemetry_task_id = ""
        self.telemetry_planned_count = 0
        self.telemetry_success_by_account = {}
        self.telemetry_failed_accounts = set()
        self.main_window_revealed = False
        self.post_login_jobs_started = False
        self.tab_scroll_canvases = []
        self.tab_scroll_areas = []
        self.browser_tree_resize_active = False
        self.authorization_window = None
        self._configure_style()
        self.build()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(250, self.start_authorization_restore)

    def restore_browser_runtime_state(self):
        """软件重启后恢复未完成队列中每个账号的累计进度与最后失败原因。"""
        queue_state = read_state(self.cfg)
        if not parse_bool(queue_state.get("queue_run_active"), False):
            return
        queue_run_id = str(queue_state.get("queue_run_id") or "").strip()
        if not queue_run_id:
            return
        completed_ids = {
            str(item) for item in (queue_state.get("queue_completed_account_ids") or []) if str(item)
        }
        labels = {
            "error": "失败",
            "logged_out": "掉号已跳过",
            "daily_limit": "今日上限已跳过",
        }
        for account in self.browser_accounts:
            account_id = str(account.get("id") or "")
            quota = max(1, _nonnegative_state_int(account.get("posts_per_run"), 1))
            account_cfg = config_for_browser_account(self.cfg, account)
            account_state = read_state(account_cfg)
            progress = account_run_resume_progress(account_state, queue_run_id, quota)
            self.browser_success_counts[account_id] = progress
            if account_id in completed_ids or progress >= quota:
                self.browser_statuses[account_id] = f"配额完成 {progress}/{quota}"
                continue
            outcome = str(account_state.get("last_account_outcome") or "")
            message = str(account_state.get("last_account_message") or "").strip()
            if outcome in labels:
                status = labels[outcome]
                if message:
                    status += f" · {message}"
                self.browser_statuses[account_id] = f"{status} · 进度 {progress}/{quota} 已保留"
            elif progress:
                self.browser_statuses[account_id] = f"进度已恢复 {progress}/{quota}"

    def _configure_style(self):
        style = ttk.Style(self.root)
        try:
            themes = set(style.theme_names())
            # Windows 的 vista 主题复选框使用标准“打勾”指示器；clam 在部分 Tk 版本中
            # 会显示成 X，容易让“已勾选”看起来像“取消/错误”。
            preferred_theme = "vista" if os.name == "nt" and "vista" in themes else "clam"
            if preferred_theme in themes:
                style.theme_use(preferred_theme)
        except Exception:
            pass
        style.configure("TFrame", background="#FFFFFF")
        style.configure("Card.TFrame", background="#FFFFFF")
        style.configure("TLabel", background="#FFFFFF", foreground="#101828", font=("Microsoft YaHei UI", 11))
        style.configure("Title.TLabel", background="#FFFFFF", foreground="#101828", font=("Microsoft YaHei UI", 14, "bold"))
        style.configure("Muted.TLabel", background="#FFFFFF", foreground="#475467", font=("Microsoft YaHei UI", 11))
        style.configure("TButton", foreground="#101828", font=("Microsoft YaHei UI", 11, "bold"), padding=(13, 8))
        style.map("TButton", foreground=[("disabled", "#101828"), ("!disabled", "#101828")])
        style.configure("TCheckbutton", background="#FFFFFF", foreground="#101828", font=("Microsoft YaHei UI", 11, "bold"), padding=(2, 3))
        style.configure("TEntry", font=("Microsoft YaHei UI", 11), foreground="#101828", padding=5)
        style.configure("TCombobox", font=("Microsoft YaHei UI", 11), padding=4)
        style.configure("Accent.TButton", background="#2563EB", foreground="#FFFFFF", borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#1D4ED8"), ("disabled", "#93C5FD")])
        style.configure("Success.TButton", background="#059669", foreground="#FFFFFF", borderwidth=0)
        style.map("Success.TButton", background=[("active", "#047857")])
        style.configure("Danger.TButton", background="#DC2626", foreground="#FFFFFF", borderwidth=0)
        style.map("Danger.TButton", background=[("active", "#B91C1C")])
        style.configure("TNotebook", background="#EEF2F7", borderwidth=0, tabmargins=(0, 0, 0, 0))
        style.configure(
            "TNotebook.Tab",
            font=("Microsoft YaHei UI", 12, "bold"),
            padding=(22, 11),
            borderwidth=0,
            relief="flat",
            background="#E5E7EB",
            foreground="#1D2939",
        )
        style.map(
            "TNotebook.Tab",
            padding=[("selected", (22, 10)), ("!selected", (22, 10))],
            background=[("selected", "#EAF2FF"), ("!selected", "#E5E7EB")],
            foreground=[("selected", "#1849A9"), ("!selected", "#1D2939")],
            font=[("selected", ("Microsoft YaHei UI", 12, "bold")), ("!selected", ("Microsoft YaHei UI", 12, "bold"))],
        )
        style.configure("Treeview", rowheight=34, font=("Microsoft YaHei UI", 11), foreground="#101828", background="#FFFFFF", fieldbackground="#FFFFFF")
        style.map("Treeview", background=[("selected", "#175CD3")], foreground=[("selected", "#FFFFFF")])
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 11, "bold"), foreground="#101828", background="#EAECF0", padding=(6, 7))

    def _card(self, parent, title, subtitle=""):
        card = ttk.Frame(parent, style="Card.TFrame", padding=14)
        ttk.Label(card, text=title, style="Title.TLabel").pack(anchor="w")
        if subtitle:
            ttk.Label(card, text=subtitle, style="Muted.TLabel").pack(anchor="w", pady=(2, 10))
        body = ttk.Frame(card)
        body.pack(fill="both", expand=True, pady=(8 if not subtitle else 0, 0))
        return card, body

    def add_row(self, parent, row, label, key, browse=None, width=74):
        ttk.Label(parent, text=label, width=20, anchor="e").grid(row=row, column=0, padx=(0, 10), pady=5, sticky="e")
        value = self.vars.get(key)
        if value is None:
            value = tk.StringVar(value=str(self.cfg.get(key, "")))
            self.vars[key] = value
        ttk.Entry(parent, textvariable=value, width=width).grid(row=row, column=1, padx=4, pady=5, sticky="we")
        if browse == "file":
            ttk.Button(parent, text="选择文件", command=lambda: self.choose_file(key)).grid(row=row, column=2, padx=(8, 0))
        elif browse == "folder":
            ttk.Button(parent, text="选择文件夹", command=lambda: self.choose_folder(key)).grid(row=row, column=2, padx=(8, 0))
        return row + 1

    def _create_scrollable_tab(self, notebook, title):
        """创建宽度自适应、内容可纵向滚动的标签页。"""
        tab = ttk.Frame(notebook)
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        notebook.add(tab, text=title)

        canvas = tk.Canvas(
            tab,
            background="#FFFFFF",
            borderwidth=0,
            highlightthickness=0,
        )
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        content = ttk.Frame(canvas, padding=10)
        content_window = canvas.create_window((0, 0), window=content, anchor="nw")

        def update_scroll_region(_event=None):
            bounds = canvas.bbox("all")
            if bounds:
                canvas.configure(scrollregion=bounds)

        def fit_content_width(event):
            canvas.itemconfigure(content_window, width=max(1, event.width))

        content.bind("<Configure>", update_scroll_region, add="+")
        canvas.bind("<Configure>", fit_content_width, add="+")
        self.tab_scroll_canvases.append(canvas)
        self.tab_scroll_areas.append((content, canvas))
        return content

    def _bind_independent_tab_scroll(self, content, canvas):
        """为单个标签绑定自己的滚轮事件，不共享滚动状态或目标。"""
        def scroll_tab(event):
            delta = int(getattr(event, "delta", 0) or 0)
            if delta:
                units = -max(1, abs(delta) // 120) if delta > 0 else max(1, abs(delta) // 120)
                canvas.yview_scroll(units, "units")
                return "break"

        stack = [content, canvas]
        while stack:
            widget = stack.pop()
            if widget.winfo_class() not in {"Treeview", "TCombobox", "Text", "Listbox"}:
                widget.bind("<MouseWheel>", scroll_tab, add="+")
            stack.extend(widget.winfo_children())

    def _show_main_window(self):
        """授权成功后才显示配置主窗口，并启动主界面后台任务。"""
        if not self.root.winfo_exists():
            return
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.main_window_revealed = True
        if self.post_login_jobs_started:
            return
        self.post_login_jobs_started = True
        self.root.after(1000, self.poll_weekly_scheduler)
        try:
            self.root.after(1500, self.auto_check_update)
        except Exception:
            pass

    def build(self):
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        header = tk.Frame(self.root, bg="#172033", height=72)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Label(header, text="抖音智能发布中心", bg="#172033", fg="#FFFFFF", font=("Microsoft YaHei UI", 18, "bold")).pack(side="left", padx=(22, 8), pady=18)
        tk.Label(header, text=f"v{APP_VERSION} · 多浏览器顺序发布", bg="#172033", fg="#D0D5DD", font=("Microsoft YaHei UI", 11, "bold")).pack(side="left", pady=24)
        tk.Button(header, text="检查更新", command=self.check_update_click, bg="#2563EB", fg="#FFFFFF", activebackground="#1D4ED8", activeforeground="#FFFFFF", relief="flat", padx=16, pady=7).pack(side="right", padx=20, pady=17)
        self.auth_button = tk.Button(
            header,
            textvariable=self.auth_button_text_var,
            command=self.authorization_button_click,
            bg="#0E9384",
            fg="#FFFFFF",
            activebackground="#107569",
            activeforeground="#FFFFFF",
            relief="flat",
            padx=14,
            pady=7,
        )
        self.auth_button.pack(side="right", padx=(4, 8), pady=17)
        tk.Label(header, textvariable=self.auth_status_var, bg="#172033", fg="#6CE9A6", font=("Microsoft YaHei UI", 11, "bold")).pack(side="right", padx=(8, 6), pady=24)
        tk.Label(header, textvariable=self.status_var, bg="#172033", fg="#D1FAE5", font=("Microsoft YaHei UI", 11, "bold")).pack(side="right", pady=24)

        notebook = ttk.Notebook(self.root)
        self.notebook = notebook
        notebook.grid(row=1, column=0, sticky="nsew", padx=14, pady=(12, 14))
        publish_tab = self._create_scrollable_tab(notebook, "发布中心")
        settings_tab = self._create_scrollable_tab(notebook, "账号与排期")
        weekly_tab = self._create_scrollable_tab(notebook, "自动启动")
        system_tab = self._create_scrollable_tab(notebook, "系统与日志")
        self._build_publish_tab(publish_tab)
        self._build_settings_tab(settings_tab)
        self._build_weekly_tab(weekly_tab)
        self._build_system_tab(system_tab)
        for content, canvas in self.tab_scroll_areas:
            self._bind_independent_tab_scroll(content, canvas)
        self._build_runtime_panel(self.root)
        self.refresh_browser_tree()
        self.write_ui(f"{APP_NAME} 已启动。已加载 {len(self.browser_accounts)} 个浏览器配置。\n")
        if self.debug_cleanup_result.get("cleaned"):
            self.write_ui(
                f"调试截图定期清理完成：本次已清理 {self.debug_cleanup_result.get('removed', 0)} 张旧截图。\n"
            )
        elif self.debug_cleanup_result.get("error"):
            self.write_ui(f"调试截图定期清理失败：{self.debug_cleanup_result['error']}\n")
        if parse_bool(self.cfg.get("rotate_accounts_each_post"), False):
            batch = normalize_rotate_batch_size(self.cfg.get("rotate_batch_size", 1))
            self.write_ui(f"当前为轮换发布：每个账号连续成功发布 {batch} 条后切换下一个；每个浏览器账号仅首次轮到时打开一次，后续直接复用，软件不主动关闭浏览器。\n")
        else:
            self.write_ui("任务按浏览器列表顺序执行；掉号账号会标记并跳过，全部浏览器处理后任务才结束。\n")

    def _build_publish_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        wait_card, waits = self._card(parent, "等待与检测")
        wait_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        fields = [
            ("每步最小等待秒", "wait_min_seconds"), ("每步最大等待秒", "wait_max_seconds"),
            ("发布间隔最小秒", "publish_interval_min_seconds"), ("发布间隔最大秒", "publish_interval_max_seconds"),
            ("上传检测间隔秒", "upload_check_interval_seconds"), ("上传最大等待秒", "upload_max_wait_seconds"),
            ("话题识别等待秒", "topic_wait_seconds"), ("创作者中心等待秒", "creator_center_wait_seconds"),
            ("单步骤失败后重试次数", "retry_times"),
            ("总体流程失败后重试次数", "overall_retry_times"),
            ("今日上限切换等待秒", "daily_limit_switch_wait_seconds"),
        ]
        self._build_field_grid(waits, fields, columns=4)

        option_card, options = self._card(parent, "运行选项")
        option_card.grid(row=1, column=0, sticky="ew")
        option_items = [
            ("发布成功后删除图片", "delete_image_after_success"),
            ("发布成功后删除文案", "delete_copy_after_success"),
            ("随机使用图片", "random_image"),
            ("复用已存在调试端口", "reuse_existing_cdp"),
            ("启动前关闭 Chrome 残留", "close_chrome_before_start"),
            ("后台运行不拉起浏览器", "no_raise_browser"),
        ]
        for index, (text, key) in enumerate(option_items):
            value = tk.BooleanVar(value=parse_bool(self.cfg.get(key), False))
            self.bool_vars[key] = value
            ttk.Checkbutton(options, text=text, variable=value).grid(
                row=index // 3, column=index % 3, padx=12, pady=8, sticky="w"
            )

        rotation_options = ttk.Frame(options)
        rotation_options.grid(row=2, column=0, columnspan=3, padx=12, pady=(10, 4), sticky="w")
        rotate_var = tk.BooleanVar(value=parse_bool(self.cfg.get("rotate_accounts_each_post"), False))
        self.bool_vars["rotate_accounts_each_post"] = rotate_var
        ttk.Checkbutton(
            rotation_options,
            text="轮换发布",
            variable=rotate_var,
        ).pack(side="left")
        ttk.Label(
            rotation_options, text="每个账号成功发布",
            font=("Microsoft YaHei UI", 11, "bold"),
        ).pack(side="left", padx=(18, 6))
        rotate_batch_var = tk.StringVar(value=str(normalize_rotate_batch_size(self.cfg.get("rotate_batch_size", 1))))
        self.vars["rotate_batch_size"] = rotate_batch_var
        self.rotate_batch_spinbox = ttk.Spinbox(
            rotation_options, from_=1, to=999,
            textvariable=rotate_batch_var, width=6,
        )
        self.rotate_batch_spinbox.pack(side="left")
        ttk.Label(
            rotation_options,
            text="条后切换下一个，循环直到各账号“发布条数”完成",
            font=("Microsoft YaHei UI", 11, "bold"),
        ).pack(side="left", padx=(6, 0))

        def sync_rotation_batch_state(*_args):
            self.rotate_batch_spinbox.configure(state="normal" if rotate_var.get() else "disabled")

        rotate_var.trace_add("write", sync_rotation_batch_state)
        sync_rotation_batch_state()

    def _build_browser_queue_card(self, parent):
        card, body = self._card(parent, "浏览器账号队列")
        tree_frame = ttk.Frame(body)
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        columns = BROWSER_QUEUE_COLUMNS
        self.browser_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=8, selectmode="browse")
        widths = normalize_browser_queue_column_widths(self.cfg.get("browser_queue_column_widths"))
        self.cfg["browser_queue_column_widths"] = dict(widths)
        for key in columns:
            self.browser_tree.heading(key, text=BROWSER_QUEUE_COLUMN_HEADINGS[key])
            self.browser_tree.column(
                key,
                width=widths[key],
                minwidth=BROWSER_QUEUE_MIN_WIDTHS[key],
                anchor="center" if key not in {"path", "excel", "schedule"} else "w",
                stretch=False,
            )
        vertical_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.browser_tree.yview)
        horizontal_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.browser_tree.xview)
        self.browser_tree.configure(
            yscrollcommand=vertical_scroll.set,
            xscrollcommand=horizontal_scroll.set,
        )
        self.browser_tree.grid(row=0, column=0, sticky="nsew")
        vertical_scroll.grid(row=0, column=1, sticky="ns")
        horizontal_scroll.grid(row=1, column=0, sticky="ew")
        self.browser_tree.bind("<Button-1>", self.on_browser_tree_click, add="+")
        self.browser_tree.bind("<ButtonPress-1>", self.on_browser_tree_resize_press, add="+")
        self.browser_tree.bind("<ButtonRelease-1>", self.on_browser_tree_resize_release, add="+")
        self.browser_tree.bind("<Double-1>", self.on_browser_tree_double_click)
        self.browser_tree.tag_configure("ok", foreground="#047857")
        self.browser_tree.tag_configure("warn", foreground="#B45309")
        self.browser_tree.tag_configure("error", foreground="#B42318")
        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(12, 0))
        for text, command in [
            ("添加", self.add_browser_account), ("复制", self.duplicate_browser_account),
            ("编辑", self.edit_browser_account),
            ("删除", self.remove_browser_account), ("上移", lambda: self.move_browser_account(-1)),
            ("下移", lambda: self.move_browser_account(1)),
            ("全选发布", lambda: self.set_all_browser_accounts_enabled(True)),
            ("全部取消", lambda: self.set_all_browser_accounts_enabled(False)),
        ]:
            ttk.Button(actions, text=text, command=command).pack(side="left", padx=(0, 6))
        table_actions = ttk.Frame(body)
        table_actions.pack(fill="x", pady=(8, 0))
        for text, command in [
            ("批量导入快捷方式", self.import_browser_shortcuts),
            ("导入队列全表", self.import_browser_queue_table),
            ("导出队列全表", self.export_browser_queue_table),
        ]:
            ttk.Button(table_actions, text=text, command=command).pack(side="left", padx=(0, 6))
        return card

    def _build_runtime_panel(self, parent):
        runtime_card, runtime = self._card(parent, "运行日志")
        runtime_card.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 14))
        runtime.grid_columnconfigure(1, weight=1)
        runtime.grid_rowconfigure(2, weight=1)

        controls = ttk.Frame(runtime)
        self.task_controls_frame = controls
        controls.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        ttk.Label(controls, text="任务控制", style="Title.TLabel").pack(side="left", padx=(0, 14))
        ttk.Button(controls, text="保存配置", command=lambda: self.save_from_ui(show_message=True)).pack(side="left", padx=(0, 6))
        ttk.Button(controls, text="打开选中浏览器", style="Success.TButton", command=self.open_browser_window).pack(side="left", padx=(0, 6))
        ttk.Button(controls, text="检查全部账号", command=self.probe_all_browsers).pack(side="left", padx=(0, 6))
        self.start_button = ttk.Button(controls, text="开始本轮发布", style="Accent.TButton", command=self.start)
        self.start_button.pack(side="left", padx=(0, 6))
        self.start_button.configure(state="disabled")
        self.retry_failed_button = ttk.Button(
            controls, text="重试勾选失败账号", command=self.retry_selected_failed_accounts,
        )
        self.retry_failed_button.pack(side="left", padx=(0, 6))
        self.retry_failed_button.configure(state="disabled")
        self.pause_button = ttk.Button(controls, text="暂停 / 继续", command=self.toggle_pause)
        self.pause_button.pack(side="left", padx=(0, 6))
        self.stop_button = ttk.Button(
            controls, text="停止任务", style="Danger.TButton", command=self.stop,
        )
        self.stop_button.pack(side="left")

        ttk.Label(runtime, textvariable=self.countdown_var, font=("Microsoft YaHei UI", 16, "bold"), foreground="#D4380D").grid(row=1, column=0, padx=(0, 18), pady=(0, 8), sticky="w")
        ttk.Label(runtime, textvariable=self.countdown_reason_var, style="Muted.TLabel").grid(row=1, column=1, pady=(0, 8), sticky="w")
        runtime_actions = ttk.Frame(runtime)
        runtime_actions.grid(row=1, column=2, pady=(0, 8), sticky="e")
        self.reset_progress_button = ttk.Button(
            runtime_actions, text="重置发布进度", command=self.reset_state,
        )
        self.reset_progress_button.pack(side="left", padx=(0, 6))
        self.clear_log_button = ttk.Button(
            runtime_actions,
            text="清空日志",
            command=lambda: self.logbox.delete("1.0", "end"),
        )
        self.clear_log_button.pack(side="left")
        self.logbox = ScrolledText(runtime, height=7, font=("Consolas", 10), bg="#0F172A", fg="#D1FAE5", insertbackground="#FFFFFF", relief="flat", padx=10, pady=8)
        self.logbox.grid(row=2, column=0, columnspan=3, sticky="nsew")

    def _build_settings_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        image_card, image_body = self._card(parent, "共享城市图片池")
        image_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        image_body.columnconfigure(1, weight=1)
        self.add_row(image_body, 0, "图片总文件夹", "image_dir", "folder")
        queue_card = self._build_browser_queue_card(parent)
        queue_card.grid(row=1, column=0, sticky="ew")

    def _build_field_grid(self, parent, fields, columns=4):
        for index, (label, key) in enumerate(fields):
            group = index % columns
            row = index // columns
            ttk.Label(parent, text=label).grid(row=row, column=group * 2, padx=(8, 5), pady=8, sticky="e")
            value = tk.StringVar(value=str(self.cfg.get(key, "")))
            self.vars[key] = value
            ttk.Entry(parent, textvariable=value, width=14).grid(row=row, column=group * 2 + 1, padx=(0, 12), pady=8, sticky="w")

    def _build_weekly_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        card, body = self._card(parent, "按星期自动启动")
        card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        enabled = tk.BooleanVar(value=parse_bool(self.cfg.get("weekly_start_enabled"), False))
        self.bool_vars["weekly_start_enabled"] = enabled
        ttk.Checkbutton(body, text="启用自动启动发布任务", variable=enabled).pack(anchor="w", pady=(2, 12))
        days_frame = ttk.Frame(body)
        days_frame.pack(fill="x", pady=8)
        selected_days = set(normalized_weekly_days(self.cfg.get("weekly_start_days", [0, 1, 2, 3, 4])))
        for index, label in enumerate(WEEKDAY_LABELS):
            value = tk.BooleanVar(value=index in selected_days)
            self.weekly_day_vars.append(value)
            ttk.Checkbutton(days_frame, text=label, variable=value).pack(side="left", padx=(0, 16))
        time_frame = ttk.Frame(body)
        time_frame.pack(fill="x", pady=(14, 8))
        ttk.Label(time_frame, text="启动时间", font=("Microsoft YaHei UI", 11, "bold")).pack(side="left")
        try:
            hour, minute = parse_clock_text(self.cfg.get("weekly_start_time", "09:00"))
        except Exception:
            hour, minute = 9, 0
        self.weekly_hour_var = tk.StringVar(value=f"{hour:02d}")
        self.weekly_minute_var = tk.StringVar(value=f"{minute:02d}")
        ttk.Combobox(time_frame, textvariable=self.weekly_hour_var, values=[f"{i:02d}" for i in range(24)], width=6, state="readonly").pack(side="left", padx=(16, 4))
        ttk.Label(time_frame, text=":", font=("Microsoft YaHei UI", 13, "bold")).pack(side="left")
        ttk.Combobox(time_frame, textvariable=self.weekly_minute_var, values=[f"{i:02d}" for i in range(60)], width=6, state="readonly").pack(side="left", padx=4)
        ttk.Button(time_frame, text="保存自动启动", style="Accent.TButton", command=lambda: self.save_from_ui(show_message=True)).pack(side="left", padx=18)
        ttk.Label(body, textvariable=self.next_schedule_var, foreground="#1D4ED8", font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", pady=(18, 6))

    def _build_system_tab(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        card, body = self._card(parent, "状态、日志与调试")
        card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        body.columnconfigure(1, weight=1)
        row = 0
        row = self.add_row(body, row, "进度文件", "state_path", "file")
        row = self.add_row(body, row, "发布日志 CSV", "log_path", "file")
        row = self.add_row(body, row, "调试截图文件夹", "debug_dir", "folder")
        self.add_row(body, row, "创作者中心地址", "creator_url")
        auth_card, auth_actions = self._card(parent, "统一授权账号")
        auth_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(auth_actions, textvariable=self.auth_status_var, font=("Microsoft YaHei UI", 11, "bold"), foreground="#067647").pack(side="left", padx=(0, 18))
        ttk.Button(auth_actions, text="登录 / 切换账号", command=lambda: self.show_authorization_login(switch_account=True)).pack(side="left", padx=6)
        ttk.Button(auth_actions, text="退出授权账号", command=self.logout_authorization).pack(side="left", padx=6)

        card, actions = self._card(parent, "程序维护")
        card.grid(row=2, column=0, sticky="ew")
        ttk.Button(actions, text="检查更新", command=self.check_update_click).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="打开程序文件夹", command=lambda: os.startfile(str(APP_DIR))).pack(side="left", padx=8)
        ttk.Button(actions, text="打开日志目录", command=self.open_log_folder).pack(side="left", padx=8)

    def start_authorization_restore(self):
        """启动时先用 DPAPI 会话自动登录；没有有效会话再显示登录入口。"""
        self.auth_status_var.set("授权状态：自动登录中")

        def task():
            error = None
            try:
                self.auth_client.auto_login()
            except Exception as exc:
                error = exc
            self.root.after(0, lambda: self.handle_authorization_restore(error))

        threading.Thread(target=task, daemon=True).start()

    def handle_authorization_restore(self, error=None):
        self.apply_authorization_state()
        if self.auth_client.state in {STATE_AUTHORIZED, STATE_GRACE}:
            self.write_ui(f"软件授权：{self.auth_client.status_text()}。\n")
            self._show_main_window()
            if self.auth_client.state == STATE_AUTHORIZED:
                self.report_client_launch()
            self.schedule_authorization_heartbeat(first_contact=True)
            return
        if error:
            self.write_ui(f"软件授权自动登录失败：{error}\n")
        else:
            self.write_ui("尚未登录软件授权账号；发布按钮已锁定。\n")
        if "--self-test" not in sys.argv:
            self.root.after(150, self.show_authorization_login)

    def apply_authorization_state(self):
        text = self.auth_client.status_text()
        if self.auth_client.state == STATE_AUTHORIZED and not self.auth_client.can_start_new_task:
            text += " · 未开通发布权益"
        self.auth_status_var.set(text)
        self.auth_button_text_var.set(
            "切换账号"
            if self.auth_client.state in {STATE_AUTHORIZED, STATE_GRACE}
            else "登录授权"
        )
        if hasattr(self, "start_button"):
            control_state = "normal" if self.auth_client.can_start_new_task else "disabled"
            self.start_button.configure(state=control_state)
            if hasattr(self, "retry_failed_button"):
                self.retry_failed_button.configure(state=control_state)

    def authorization_button_click(self):
        self.show_authorization_login(
            switch_account=self.auth_client.state in {STATE_AUTHORIZED, STATE_GRACE}
        )

    def show_authorization_login(self, switch_account=False):
        if self.proc and self.proc.poll() is None:
            messagebox.showwarning(
                "任务正在运行",
                "切换软件授权账号前必须先停止当前发布任务。",
            )
            return
        if self.authorization_window is not None:
            try:
                if self.authorization_window.winfo_exists():
                    self.authorization_window.deiconify()
                    self.authorization_window.lift()
                    self.authorization_window.focus_force()
                    return
            except Exception:
                self.authorization_window = None
        login_required = not switch_account and not self.main_window_revealed
        window = tk.Toplevel(self.root)
        self.authorization_window = window
        window.title("软件用户登录" if not switch_account else "切换软件授权账号")
        window.geometry("440x325")
        window.resizable(False, False)
        if self.main_window_revealed:
            window.transient(self.root)
        window.grab_set()
        try:
            if LOGO_ICO.exists():
                window.iconbitmap(str(LOGO_ICO))
        except Exception:
            pass
        body = ttk.Frame(window, padding=24)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="统一软件授权", font=("Microsoft YaHei UI", 18, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(body, text="软件账号", width=11, anchor="e").grid(row=1, column=0, padx=(0, 10), pady=(20, 7), sticky="e")
        try:
            saved_login = self.login_preference_store.load(AUTH_PRODUCT_CODE) or {}
        except Exception:
            saved_login = {}
        saved_account = str(saved_login.get("account") or "")
        saved_password = str(saved_login.get("password") or "")
        remember_password_default = bool(
            saved_login.get("rememberPassword") and saved_account and saved_password
        )
        remember_account_default = bool(
            saved_login.get("rememberAccount") and saved_account
        ) or remember_password_default
        account_var = tk.StringVar(value=saved_account or self.auth_client.account)
        account_entry = ttk.Entry(body, textvariable=account_var, width=31)
        account_entry.grid(row=1, column=1, pady=(20, 7), sticky="we")
        ttk.Label(body, text="密码", width=11, anchor="e").grid(row=2, column=0, padx=(0, 10), pady=7, sticky="e")
        password_var = tk.StringVar(value=saved_password if remember_password_default else "")
        password_entry = ttk.Entry(body, textvariable=password_var, width=31, show="●")
        password_entry.grid(row=2, column=1, pady=7, sticky="we")
        remember_account_var = tk.BooleanVar(value=remember_account_default)
        remember_password_var = tk.BooleanVar(value=remember_password_default)
        remember_row = ttk.Frame(body)
        remember_row.grid(row=3, column=1, sticky="w", pady=(4, 3))

        def remember_password_changed():
            if remember_password_var.get():
                remember_account_var.set(True)

        def remember_account_changed():
            if not remember_account_var.get():
                remember_password_var.set(False)

        ttk.Checkbutton(
            remember_row,
            text="记住账号",
            variable=remember_account_var,
            command=remember_account_changed,
        ).pack(side="left")
        ttk.Checkbutton(
            remember_row,
            text="记住密码",
            variable=remember_password_var,
            command=remember_password_changed,
        ).pack(side="left", padx=(18, 0))
        status_var = tk.StringVar(value="")
        ttk.Label(body, textvariable=status_var, foreground="#B42318", wraplength=375).grid(row=4, column=0, columnspan=2, sticky="w", pady=(5, 8))
        buttons = ttk.Frame(body)
        buttons.grid(row=5, column=0, columnspan=2, sticky="e")
        def close_login_window():
            try:
                window.grab_release()
            except Exception:
                pass
            self.authorization_window = None
            window.destroy()
            signed_out_after_switch = (
                switch_account and self.auth_client.state == STATE_SIGNED_OUT
            )
            if (login_required or signed_out_after_switch) and self.root.winfo_exists():
                self.root.destroy()

        window.protocol("WM_DELETE_WINDOW", close_login_window)
        ttk.Button(buttons, text="取消", command=close_login_window).pack(side="right", padx=(8, 0))
        login_button = ttk.Button(buttons, text="登录", style="Accent.TButton")
        login_button.pack(side="right")
        body.columnconfigure(1, weight=1)

        def submit(_event=None):
            account = account_var.get().strip()
            password = password_var.get()
            if not account or not password:
                status_var.set("请输入软件账号和密码。")
                return
            login_button.configure(state="disabled")
            status_var.set("正在连接统一授权后台……")

            def login_task():
                try:
                    if switch_account and self.auth_client.state != STATE_SIGNED_OUT:
                        self.auth_client.logout()
                    self.auth_client.login(account, password)
                    self.root.after(0, login_success)
                except Exception as exc:
                    self.root.after(0, lambda exc=exc: login_failed(exc))

            threading.Thread(target=login_task, daemon=True).start()

        def login_success():
            keep_password = bool(remember_password_var.get())
            keep_account = bool(remember_account_var.get()) or keep_password
            try:
                if keep_account:
                    self.login_preference_store.save(
                        {
                            "productCode": AUTH_PRODUCT_CODE,
                            "rememberAccount": True,
                            "rememberPassword": keep_password,
                            "account": account_var.get().strip(),
                            "password": password_var.get() if keep_password else "",
                        }
                    )
                else:
                    self.login_preference_store.delete()
            except Exception as exc:
                self.write_ui(f"登录偏好未能保存：{exc}\n")
            self.auth_force_logout_prompt_shown = False
            self.apply_authorization_state()
            self.write_ui(f"软件授权登录成功：{self.auth_client.account}。\n")
            self.authorization_window = None
            window.destroy()
            self._show_main_window()
            self.report_client_launch()
            self.schedule_authorization_heartbeat(first_contact=True)

        def login_failed(exc):
            self.apply_authorization_state()
            login_button.configure(state="normal")
            status_var.set(str(exc))
            password_var.set("")
            password_entry.focus_set()

        login_button.configure(command=submit)
        window.bind("<Return>", submit)
        (password_entry if account_var.get() else account_entry).focus_set()
        window.update_idletasks()
        if login_required:
            width = window.winfo_width()
            height = window.winfo_height()
            x = max(0, (window.winfo_screenwidth() - width) // 2)
            y = max(0, (window.winfo_screenheight() - height) // 2)
            window.geometry(f"{width}x{height}+{x}+{y}")
        window.deiconify()
        window.lift()
        window.focus_force()

    def logout_authorization(self):
        if self.auth_client.state == STATE_SIGNED_OUT:
            messagebox.showinfo("软件授权", "当前尚未登录软件授权账号。")
            return
        if self.proc and self.proc.poll() is None:
            if not messagebox.askyesno("退出授权", "退出授权会停止当前发布任务，是否继续？"):
                return
            self.stop()
        if not messagebox.askyesno("退出授权", "确定退出当前软件授权账号吗？"):
            return

        def task():
            self.auth_client.logout()
            self.root.after(0, self.handle_authorization_logout)

        threading.Thread(target=task, daemon=True).start()

    def handle_authorization_logout(self):
        self.apply_authorization_state()
        self.write_ui("已退出软件授权账号，本机 DPAPI 会话已清除。\n")
        self.main_window_revealed = False
        self.root.withdraw()
        self.root.after(100, lambda: self.show_authorization_login(switch_account=False))

    def schedule_authorization_heartbeat(self, first_contact=False):
        delay = self.auth_client.next_heartbeat_delay(
            first_contact=first_contact,
            configured_seconds=AUTH_HEARTBEAT_SECONDS,
        )
        self.root.after(max(1000, int(delay * 1000)), self.run_authorization_heartbeat)

    def run_authorization_heartbeat(self):
        if self.auth_heartbeat_inflight:
            return
        if self.auth_client.state not in {STATE_AUTHORIZED, STATE_GRACE}:
            return
        self.auth_heartbeat_inflight = True

        def task():
            error = None
            try:
                # 与 ByxxPublisher 一致：短暂网络故障在进入下个周期前最多快速重试 3 次。
                for attempt in range(3):
                    self.auth_client.heartbeat()
                    if self.auth_client.state == STATE_AUTHORIZED:
                        break
                    if attempt < 2:
                        time.sleep(0.5 if attempt == 0 else 2.0)
            except Exception as exc:
                error = exc
            self.root.after(0, lambda: self.handle_authorization_heartbeat(error))

        threading.Thread(target=task, daemon=True).start()

    def handle_authorization_heartbeat(self, error=None):
        self.auth_heartbeat_inflight = False
        previous = self.auth_status_var.get()
        self.apply_authorization_state()
        current = self.auth_status_var.get()
        if current != previous:
            self.write_ui(f"软件授权状态变化：{current}。\n")
        if self.auth_client.state == STATE_FORCE_LOGOUT:
            self.handle_authorization_revoked(error)
            return
        if error and self.auth_client.state != STATE_GRACE:
            self.write_ui(f"授权心跳将在后台重试：{error}\n")
        self.schedule_authorization_heartbeat(first_contact=False)

    def handle_authorization_revoked(self, error=None):
        if self.proc and self.proc.poll() is None:
            self.stop()
        self.write_ui(f"在线授权已失效，任务已停止：{error or '后台已撤销当前会话'}。\n")
        if self.auth_force_logout_prompt_shown:
            return
        self.auth_force_logout_prompt_shown = True
        should_login = messagebox.askyesno(
            "软件授权已失效",
            "后台授权已暂停、撤销、到期或当前会话已失效。\n\n"
            "正在运行的发布任务已经停止。是否现在重新登录软件账号？",
            icon="warning",
        )
        if should_login:
            self.main_window_revealed = False
            self.root.withdraw()
            self.show_authorization_login(switch_account=False)
        else:
            self.root.destroy()

    def send_telemetry_async(self, operation, description="遥测"):
        """遥测失败不能阻塞发布；只记录脱敏错误，不上报路径或正文。"""
        if self.auth_client.state != STATE_AUTHORIZED:
            return

        def task():
            try:
                operation()
            except Exception as exc:
                self.root.after(
                    0,
                    lambda exc=exc: self.write_ui(f"{description}暂未送达，将由后续事件继续统计：{exc}\n"),
                )

        threading.Thread(target=task, daemon=True).start()

    def report_client_launch(self):
        session_id = str((self.auth_client.session or {}).get("sessionId") or uuid.uuid4().hex)
        self.send_telemetry_async(
            lambda: self.auth_client.report_usage(
                "app.launch",
                1,
                f"app.launch:{session_id}:{APP_VERSION}",
            ),
            "启动遥测",
        )

    def reset_task_telemetry(self):
        self.telemetry_task_id = uuid.uuid4().hex
        accounts = browser_accounts_from_config(self.cfg, enabled_only=True)
        planned = sum(max(1, int(item.get("posts_per_run", 1) or 1)) for item in accounts)
        self.telemetry_planned_count = planned
        self.telemetry_success_by_account = {}
        self.telemetry_failed_accounts = set()

    def finish_task_telemetry(self, status):
        if not self.telemetry_task_id:
            return
        task_id = self.telemetry_task_id
        succeeded = sum(self.telemetry_success_by_account.values())
        failed = len(self.telemetry_failed_accounts)
        attempted = max(succeeded + failed, min(self.telemetry_planned_count, succeeded + failed))
        attempted = max(attempted, succeeded + failed)
        final_status = str(status or "stopped")
        self.send_telemetry_async(
            lambda: self.auth_client.report_task(
                final_status,
                attempted,
                succeeded,
                failed,
                task_type="publish.batch",
                idempotency_key=f"publish.batch:{task_id}:{final_status}",
            ),
            "任务遥测",
        )
        if succeeded:
            self.send_telemetry_async(
                lambda: self.auth_client.report_usage(
                    "publish.success",
                    succeeded,
                    f"publish.success:{task_id}",
                ),
                "成功用量遥测",
            )
        self.telemetry_task_id = ""

    def auto_check_update(self):
        self.check_update_async(auto=True)

    def check_update_click(self):
        self.check_update_async(auto=False)

    def check_update_async(self, auto=False):
        def task():
            try:
                info = available_update_info(self.auth_client)
                self.root.after(0, lambda: self.handle_update_result(info, auto))
            except Exception as exc:
                if not auto:
                    self.root.after(0, lambda: messagebox.showerror("检查更新失败", str(exc)))
        threading.Thread(target=task, daemon=True).start()

    def handle_update_result(self, info, auto=False):
        if not info:
            if not auto:
                messagebox.showinfo("检查更新", f"当前已经是最新版本：v{APP_VERSION}")
            return
        latest = str(info.get("latest_version") or info.get("version") or "")
        notes = format_release_notes(info)
        prompt = f"发现新版本：v{APP_VERSION} → v{latest}\n\n"
        if notes:
            prompt += notes + "\n\n"
        source_names = {
            "unified-platform": "统一管理后台",
            "github-raw": "GitHub raw 备用源",
            "jsdelivr": "jsDelivr 备用源",
        }
        prompt += f"更新策略来源：{source_names.get(info.get('source'), info.get('source') or '未知')}。\n"
        prompt += "是否现在下载并更新？本机配置和 DPAPI 授权会话不会放入安装目录，也不会被覆盖。"
        if messagebox.askyesno("发现新版本", prompt):
            self.start_online_update(info)

    def start_online_update(self, info):
        if self.proc and self.proc.poll() is None:
            messagebox.showwarning("正在运行", "发布任务正在运行，请先停止任务后再更新。")
            return
        self.save_from_ui(show_message=False)
        updater = (
            INSTALL_DIR / "DouyinPublisherUpdater.exe"
            if IS_FROZEN
            else SOURCE_APP_DIR / "online_updater.py"
        )
        if not updater.exists():
            messagebox.showerror("无法更新", f"缺少更新器文件：{updater}")
            return
        preserve = info.get("preserve_paths") or PRESERVE_UPDATE_PATHS
        preserve_arg = json.dumps(preserve if isinstance(preserve, list) else PRESERVE_UPDATE_PATHS, ensure_ascii=False)
        updater_command = []
        if IS_FROZEN:
            # 更新器复制到临时目录运行，避免 Windows 锁住安装目录内的更新器 EXE。
            temporary_updater = Path(tempfile.gettempdir()) / f"DouyinPublisherUpdater-{uuid.uuid4().hex}.exe"
            shutil.copy2(updater, temporary_updater)
            updater_command = [str(temporary_updater)]
        else:
            updater_command = [worker_python_executable(), str(updater)]
        cmd = updater_command + [
            "--install-dir", str(INSTALL_DIR),
            "--current-pid", str(os.getpid()),
            "--download-url", str(info.get("download_url", "")),
            "--sha256", str(info.get("sha256", "")),
            "--launch", str(sys.executable if IS_FROZEN else Path(__file__).resolve()),
            "--preserve", preserve_arg,
        ]
        try:
            subprocess.Popen(cmd, cwd=str(INSTALL_DIR), **hidden_subprocess_kwargs())
            messagebox.showinfo("开始更新", "更新器已启动。当前窗口会关闭，更新完成后会自动重新打开。")
            self.root.after(300, self.root.destroy)
        except Exception as exc:
            messagebox.showerror("启动更新失败", repr(exc))

    def choose_file(self, key):
        if key == "excel_path":
            path = filedialog.askopenfilename(
                title="选择文案 Excel",
                filetypes=[("Excel 文案表", "*.xlsx *.xls *.xlsm")],
            )
        else:
            path = filedialog.askopenfilename()
        if path:
            self.vars[key].set(path)

    def choose_folder(self, key):
        path = filedialog.askdirectory()
        if path:
            self.vars[key].set(path)

    def collect_cfg(self, allow_no_enabled=False):
        cfg = self.cfg.copy()
        int_keys = {
            "retry_times", "overall_retry_times",
            "rotate_batch_size", "daily_limit_switch_wait_seconds",
        }
        float_keys = {
            "wait_min_seconds", "wait_max_seconds", "publish_interval_min_seconds",
            "publish_interval_max_seconds", "upload_check_interval_seconds",
            "upload_max_wait_seconds", "topic_wait_seconds", "creator_center_wait_seconds",
        }
        for key, value in self.vars.items():
            text_value = value.get().strip()
            if key in int_keys:
                cfg[key] = int(text_value or 0)
            elif key in float_keys:
                cfg[key] = float(text_value or 0)
            else:
                cfg[key] = text_value
        for key, value in self.bool_vars.items():
            cfg[key] = bool(value.get())
        cfg["rotate_batch_size"] = normalize_rotate_batch_size(cfg.get("rotate_batch_size", 1))
        cfg["retry_times"] = normalize_retry_count(cfg.get("retry_times", 3), 3)
        cfg["overall_retry_times"] = normalize_retry_count(cfg.get("overall_retry_times", 1), 1)
        cfg["daily_limit_switch_wait_seconds"] = normalize_daily_limit_switch_wait_seconds(
            cfg.get("daily_limit_switch_wait_seconds", 10)
        )
        cfg["weekly_start_days"] = [index for index, value in enumerate(self.weekly_day_vars) if value.get()]
        cfg["weekly_start_time"] = f"{self.weekly_hour_var.get()}:{self.weekly_minute_var.get()}"
        parse_clock_text(cfg["weekly_start_time"])
        if cfg.get("weekly_start_enabled") and not cfg["weekly_start_days"]:
            raise ValueError("启用自动启动时，至少选择一个星期。")
        cfg["browser_accounts"] = [normalize_browser_account(account, index, cfg) for index, account in enumerate(self.browser_accounts)]
        enabled = [account for account in cfg["browser_accounts"] if account.get("enabled", True)]
        if not enabled and not allow_no_enabled:
            raise ValueError("至少启用一个浏览器账号。")
        if enabled:
            validate_browser_queue_ports(enabled)
        for account in enabled:
            if not account.get("browser_path"):
                raise ValueError(f"浏览器账号“{account.get('name')}”未设置快捷方式或 EXE。")
        for account in enabled:
            account_cfg = config_for_browser_account(cfg, account)
            if not str(account_cfg.get("sheet_name") or "").strip():
                raise ValueError(f"队列项“{account['name']}”的工作表名不能为空。")
            try:
                validate_publish_records_and_images(account_cfg)
            except Exception as exc:
                raise ValueError(f"队列项“{account['name']}”内容检查失败：{exc}") from exc
            if account_cfg.get("use_schedule", True):
                build_slots(account_cfg, log_result=False)
        queue_rows = enabled or cfg["browser_accounts"]
        if not queue_rows:
            raise ValueError("浏览器账号队列不能为空。")
        first = queue_rows[0]
        cfg["browser_path"] = first["browser_path"]
        cfg["browser_user_data_dir"] = first.get("browser_user_data_dir", "")
        cfg["browser_profile_directory"] = first.get("browser_profile_directory", "")
        cfg["cdp_port"] = int(first["cdp_port"])
        cfg["music_required"] = True
        cfg["keep_browser_open"] = True
        cfg["auto_switch_graphic_mode"] = True
        cfg["platform_music_only"] = True
        return cfg

    def save_from_ui(self, show_message=False, allow_no_enabled=False):
        try:
            self.cfg = self.collect_cfg(allow_no_enabled=allow_no_enabled)
            self.browser_accounts = browser_accounts_from_config(self.cfg)
            save_config(self.cfg)
            self.refresh_browser_tree()
            mode_text = (
                f"每账号连续发 {normalize_rotate_batch_size(self.cfg.get('rotate_batch_size', 1))} 条后轮换"
                if parse_bool(self.cfg.get("rotate_accounts_each_post"), False)
                else "单账号完成配额后再切换"
            )
            self.write_ui(f"配置已保存；下一次任务会使用当前浏览器顺序和发布配额。当前模式：{mode_text}。\n")
            if show_message:
                messagebox.showinfo("保存成功", "配置已保存。")
            return True
        except Exception as exc:
            if show_message:
                messagebox.showwarning("配置无效", str(exc))
            else:
                self.write_ui(f"配置保存失败：{exc}\n")
            return False

    def refresh_browser_tree(self):
        if not hasattr(self, "browser_tree"):
            return
        selected = self.browser_tree.selection()
        selected_id = selected[0] if selected else ""
        for item in self.browser_tree.get_children():
            self.browser_tree.delete(item)
        for index, account in enumerate(self.browser_accounts, start=1):
            account_id = str(account["id"])
            status = self.browser_statuses.get(account_id, "等待任务")
            published_count = int(self.browser_success_counts.get(account_id, 0) or 0)
            tag = "error" if any(word in status for word in ("掉号", "失败", "错误")) else ("ok" if any(word in status for word in ("完成", "就绪")) else "warn" if any(word in status for word in ("检查", "上限")) else "")
            schedule_text = "立即发布"
            if parse_bool(account.get("use_schedule"), True):
                start_date = str(account.get("publish_start_date") or "")
                end_date = str(account.get("publish_end_date") or start_date)
                date_text = start_date if start_date == end_date else f"{start_date}~{end_date}"
                schedule_text = (
                    f"{date_text} {account.get('start_hour', 0)}-{account.get('end_hour', 0)}时 "
                    f"{account.get('custom_minutes') or (str(account.get('per_hour_count', 1)) + '条/时')}"
                )
            self.browser_tree.insert(
                "", "end", iid=account_id,
                values=(
                    "✓" if account.get("enabled", True) else "□",
                    index,
                    account["name"],
                    account["browser_path"],
                    account.get("excel_path", ""),
                    schedule_text,
                    account["cdp_port"],
                    account["posts_per_run"],
                    published_count,
                    status,
                ),
                tags=(tag,) if tag else (),
            )
        if selected_id and self.browser_tree.exists(selected_id):
            self.browser_tree.selection_set(selected_id)
        elif self.browser_accounts:
            self.browser_tree.selection_set(self.browser_accounts[0]["id"])

    def selected_browser_index(self):
        selected = self.browser_tree.selection() if hasattr(self, "browser_tree") else ()
        if not selected:
            return None
        for index, account in enumerate(self.browser_accounts):
            if account["id"] == selected[0]:
                return index
        return None

    def on_browser_tree_click(self, event):
        if self.browser_tree.identify_region(event.x, event.y) != "cell":
            return None
        if self.browser_tree.identify_column(event.x) != "#1":
            return None
        account_id = self.browser_tree.identify_row(event.y)
        if account_id:
            self.toggle_browser_account_enabled(account_id)
        return "break"

    def on_browser_tree_resize_press(self, event):
        self.browser_tree_resize_active = self.browser_tree.identify_region(event.x, event.y) == "separator"

    def on_browser_tree_resize_release(self, _event):
        if not self.browser_tree_resize_active:
            return None
        self.browser_tree_resize_active = False
        self.root.after_idle(self.persist_browser_tree_column_widths)
        return None

    def persist_browser_tree_column_widths(self):
        if not hasattr(self, "browser_tree"):
            return False
        widths = normalize_browser_queue_column_widths({
            key: self.browser_tree.column(key, "width") for key in BROWSER_QUEUE_COLUMNS
        })
        if widths == normalize_browser_queue_column_widths(self.cfg.get("browser_queue_column_widths")):
            return False
        self.cfg["browser_queue_column_widths"] = dict(widths)
        try:
            save_config(self.cfg)
            return True
        except Exception as exc:
            self.write_ui(f"保存账号队列表列宽失败：{exc}\n")
            return False

    def on_browser_tree_double_click(self, event):
        if self.browser_tree.identify_column(event.x) == "#1":
            return "break"
        self.edit_browser_account()
        return None

    def toggle_browser_account_enabled(self, account_id):
        if not self.ensure_account_editable():
            return False
        account = next((item for item in self.browser_accounts if str(item.get("id")) == str(account_id)), None)
        if account is None:
            return False
        account["enabled"] = not parse_bool(account.get("enabled"), True)
        self.browser_statuses[str(account["id"])] = "等待任务" if account["enabled"] else "未启用"
        self.refresh_browser_tree()
        self.browser_tree.selection_set(str(account["id"]))
        self.browser_tree.focus(str(account["id"]))
        return True

    def set_all_browser_accounts_enabled(self, enabled):
        if not self.ensure_account_editable():
            return False
        enabled = bool(enabled)
        for account in self.browser_accounts:
            account["enabled"] = enabled
            self.browser_statuses[str(account["id"])] = "等待任务" if enabled else "未启用"
        self.refresh_browser_tree()
        return True

    def ensure_account_editable(self):
        if self.proc and self.proc.poll() is None:
            if self.proc_mode == "worker" and self.is_paused:
                return True
            messagebox.showwarning(
                "任务运行中",
                "运行中不能直接修改浏览器账号队列。请先点击“暂停 / 继续”进入暂停状态，"
                "暂停后即可添加、编辑、删除、勾选或调整队列顺序。",
            )
            return False
        return True

    def next_available_port(self):
        used = {int(account.get("cdp_port", 0) or 0) for account in self.browser_accounts}
        port = 9222
        while port in used and port < 65535:
            port += 1
        return port

    def add_browser_account(self):
        if not self.ensure_account_editable():
            return
        initial = normalize_browser_account({
            "id": new_browser_queue_id(),
            "name": f"浏览器{len(self.browser_accounts) + 1}", "cdp_port": self.next_available_port(),
            "posts_per_run": 1, "enabled": True, "close_after_finish": True,
        }, len(self.browser_accounts), self.cfg)
        dialog = BrowserAccountDialog(self.root, initial, "添加浏览器队列项")
        if dialog.result:
            dialog.result["id"] = initial["id"]
            self.browser_accounts.append(dialog.result)
            self.refresh_browser_tree()

    def duplicate_browser_account(self):
        if not self.ensure_account_editable():
            return
        index = self.selected_browser_index()
        if index is None:
            messagebox.showinfo("请选择队列项", "请先选择要复制的浏览器队列项。")
            return
        duplicate = copy.deepcopy(self.browser_accounts[index])
        duplicate["id"] = new_browser_queue_id()
        duplicate["name"] = f"{duplicate.get('name') or '浏览器'}（再次执行）"
        self.browser_accounts.insert(index + 1, duplicate)
        self.refresh_browser_tree()
        self.browser_tree.selection_set(duplicate["id"])

    def edit_browser_account(self):
        if not self.ensure_account_editable():
            return
        index = self.selected_browser_index()
        if index is None:
            messagebox.showinfo("请选择账号", "请先在列表中选择一个浏览器账号。")
            return
        dialog = BrowserAccountDialog(self.root, self.browser_accounts[index], "编辑浏览器账号")
        if dialog.result:
            self.browser_accounts[index] = dialog.result
            self.refresh_browser_tree()

    def remove_browser_account(self):
        if not self.ensure_account_editable():
            return
        index = self.selected_browser_index()
        if index is None:
            return
        account = self.browser_accounts[index]
        if messagebox.askyesno("删除浏览器账号", f"确定从任务队列删除“{account['name']}”吗？\n不会删除浏览器用户目录。"):
            self.browser_accounts.pop(index)
            self.refresh_browser_tree()

    def move_browser_account(self, offset):
        if not self.ensure_account_editable():
            return
        index = self.selected_browser_index()
        if index is None:
            return
        target = index + int(offset)
        if not 0 <= target < len(self.browser_accounts):
            return
        self.browser_accounts[index], self.browser_accounts[target] = self.browser_accounts[target], self.browser_accounts[index]
        account_id = self.browser_accounts[target]["id"]
        self.refresh_browser_tree()
        self.browser_tree.selection_set(account_id)

    def import_browser_shortcuts(self):
        if not self.ensure_account_editable():
            return
        paths = filedialog.askopenfilenames(
            title="选择多个浏览器快捷方式",
            initialdir="D:\\" if Path("D:\\").exists() else None,
            filetypes=[("浏览器快捷方式", "*.lnk"), ("浏览器程序", "*.exe")],
        )
        added = 0
        for path in paths:
            existing = next(
                (item for item in self.browser_accounts if os.path.normcase(item.get("browser_path", "")) == os.path.normcase(path)),
                None,
            )
            if existing:
                account = copy.deepcopy(existing)
                account["name"] = f"{existing.get('name') or Path(path).stem}（再次执行）"
            else:
                account = normalize_browser_account({
                    "name": Path(path).stem, "browser_path": path,
                    "cdp_port": self.next_available_port(), "posts_per_run": 1,
                    "enabled": True, "close_after_finish": True,
                }, len(self.browser_accounts), self.cfg)
            account["id"] = new_browser_queue_id()
            self.browser_accounts.append(account)
            added += 1
        self.refresh_browser_tree()
        if paths:
            self.write_ui(f"批量导入完成：新增 {added} 个队列项；重复浏览器已保留。\n")

    def export_browser_queue_table(self):
        if not self.browser_accounts:
            messagebox.showinfo("没有队列", "当前没有可导出的浏览器队列项。")
            return
        path = filedialog.asksaveasfilename(
            title="导出浏览器账号队列全表",
            defaultextension=".xlsx",
            initialfile="浏览器账号队列.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return
        workbook = None
        try:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "浏览器账号队列"
            sheet.append([label for label, _key in BROWSER_QUEUE_TABLE_COLUMNS])
            bool_keys = {"enabled", "close_after_finish", "use_schedule"}
            for order, account in enumerate(self.browser_accounts, start=1):
                row = []
                for _label, key in BROWSER_QUEUE_TABLE_COLUMNS:
                    if key == "_order":
                        value = order
                    else:
                        value = account.get(key, "")
                    if key in bool_keys:
                        value = "是" if parse_bool(value, False) else "否"
                    row.append(value)
                sheet.append(row)
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for column in sheet.columns:
                max_width = max(len(str(cell.value or "")) for cell in column)
                sheet.column_dimensions[column[0].column_letter].width = min(45, max(10, max_width + 2))
            workbook.save(path)
            workbook.close()
            workbook = None
            self.write_ui(f"已导出浏览器账号队列全表：{path}\n")
            messagebox.showinfo("导出完成", f"已导出 {len(self.browser_accounts)} 个队列项。")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
        finally:
            try:
                if workbook is not None:
                    workbook.close()
            except Exception:
                pass

    def import_browser_queue_table(self):
        if not self.ensure_account_editable():
            return
        path = filedialog.askopenfilename(
            title="导入浏览器账号队列全表",
            filetypes=[("Excel 队列表", "*.xlsx *.xlsm")],
        )
        if not path:
            return
        workbook = None
        try:
            workbook = load_workbook(path, data_only=True, read_only=True)
            sheet = workbook[workbook.sheetnames[0]]
            headers = {
                str(cell.value or "").strip(): index
                for index, cell in enumerate(sheet[1])
                if str(cell.value or "").strip()
            }
            required = {"账号名称", "浏览器路径", "CDP端口", "发布条数", "文案Excel", "工作表名"}
            missing = sorted(required - set(headers))
            if missing:
                raise ValueError("队列表缺少列：" + "、".join(missing))
            imported = []
            bool_keys = {"enabled", "close_after_finish", "use_schedule"}
            int_keys = {"cdp_port", "posts_per_run", "start_hour", "end_hour", "per_hour_count"}
            for row_number, cells in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                if not any(value not in {None, ""} for value in cells):
                    continue
                raw = {}
                for label, key in BROWSER_QUEUE_TABLE_COLUMNS:
                    if key == "_order" or label not in headers:
                        continue
                    column = headers[label]
                    value = cells[column] if column < len(cells) else ""
                    if key in bool_keys:
                        value = parse_bool(value, True)
                    elif key in int_keys:
                        value = int(value or 0)
                    raw[key] = value
                if not str(raw.get("id") or "").strip():
                    raw["id"] = new_browser_queue_id()
                account = normalize_browser_account(raw, len(imported), self.cfg)
                if not account.get("browser_path"):
                    raise ValueError(f"第 {row_number} 行没有浏览器路径。")
                imported.append(account)
            workbook.close()
            workbook = None
            if not imported:
                raise ValueError("队列表中没有可导入的数据行。")
            validate_browser_queue_ports([item for item in imported if item.get("enabled", True)])
            if self.browser_accounts and not messagebox.askyesno(
                "替换当前队列",
                f"将用表格中的 {len(imported)} 个队列项替换当前队列，是否继续？",
            ):
                return
            self.browser_accounts = browser_accounts_from_config(
                {**self.cfg, "browser_accounts": imported}
            )
            self.refresh_browser_tree()
            self.write_ui(f"已导入浏览器账号队列全表：{len(imported)} 个队列项；请保存配置。\n")
            messagebox.showinfo("导入完成", f"已导入 {len(imported)} 个队列项。")
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
        finally:
            try:
                if workbook is not None:
                    workbook.close()
            except Exception:
                pass

    def write_ui(self, text):
        if not hasattr(self, "logbox"):
            return
        self.logbox.insert("end", text)
        self.logbox.see("end")

    @staticmethod
    def format_countdown(seconds):
        try:
            total = max(0, int(math.ceil(float(seconds or 0))))
        except Exception:
            total = 0
        hours, rem = divmod(total, 3600)
        minutes, secs = divmod(rem, 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}" if hours else f"{minutes:02d}:{secs:02d}"

    def handle_worker_event(self, line):
        text_line = str(line or "").strip()
        if text_line.startswith(AUTO_PAUSE_EVENT_PREFIX):
            try:
                payload = json.loads(text_line[len(AUTO_PAUSE_EVENT_PREFIX):])
                message = str(payload.get("message") or "软件已自动暂停，请查看运行日志。")
            except Exception:
                message = "软件已自动暂停，请查看运行日志。"
            self.is_paused = True
            self.status_var.set("状态：自动暂停提醒")
            self.countdown_reason_var.set(message)
            self.write_ui(message + "\n")
            if self.telemetry_task_id:
                self.send_telemetry_async(
                    lambda: self.auth_client.report_error(
                        "publish.batch.paused",
                        "publish.batch",
                        retryable=True,
                        idempotency_key=f"publish.batch.paused:{self.telemetry_task_id}:{uuid.uuid4().hex[:8]}",
                    ),
                    "暂停错误遥测",
                )
            try:
                self.root.deiconify()
                self.root.lift()
            except Exception:
                pass
            messagebox.showwarning("自动暂停提醒", message, parent=self.root)
            return True
        if text_line.startswith(BROWSER_STATUS_EVENT_PREFIX):
            try:
                payload = json.loads(text_line[len(BROWSER_STATUS_EVENT_PREFIX):])
                labels = {
                    "launching": "正在启动", "checking": "检查登录", "ready": "账号就绪",
                    "publishing": "正在发布", "done": "配额完成", "logged_out": "掉号已跳过",
                    "error": "失败已跳过", "waiting": "等待任务",
                    "daily_limit": "今日上限已跳过",
                }
                account_id = str(payload.get("id", ""))
                raw_status = str(payload.get("status", ""))
                message = str(payload.get("message", "")).strip()
                status = labels.get(str(payload.get("status", "")), str(payload.get("status", "")))
                success_count = int(payload.get("success_count", 0) or 0)
                quota = int(payload.get("quota", 0) or 0)
                if quota and payload.get("status") in {"publishing", "done"}:
                    status += f" {success_count}/{quota}"
                if message and raw_status in {"error", "logged_out", "daily_limit"}:
                    status += f" · {message}"
                self.browser_statuses[account_id] = status
                self.browser_success_counts[account_id] = max(
                    success_count,
                    int(self.browser_success_counts.get(account_id, 0) or 0),
                )
                self.telemetry_success_by_account[account_id] = max(
                    success_count,
                    int(self.telemetry_success_by_account.get(account_id, 0) or 0),
                )
                if raw_status in {"error", "logged_out"}:
                    self.telemetry_failed_accounts.add(account_id or f"unknown-{len(self.telemetry_failed_accounts)}")
                self.refresh_browser_tree()
                if payload.get("name"):
                    self.status_var.set(f"状态：{payload.get('name')} · {status}")
                if message:
                    self.countdown_reason_var.set(message)
            except Exception:
                pass
            return True
        if not text_line.startswith(COUNTDOWN_EVENT_PREFIX):
            return False
        try:
            payload = json.loads(text_line[len(COUNTDOWN_EVENT_PREFIX):])
            action = str(payload.get("action") or "")
            seconds = float(payload.get("seconds") or 0)
            reason = str(payload.get("reason") or "等待下一步")
            if action in {"start", "tick", "resume"}:
                self.countdown_var.set("倒计时：" + self.format_countdown(seconds))
                self.countdown_reason_var.set(reason)
                if action == "resume":
                    self.is_paused = False
            elif action == "pause":
                self.is_paused = True
                self.status_var.set("状态：已暂停")
                self.countdown_var.set("暂停于：" + self.format_countdown(seconds))
                self.countdown_reason_var.set(reason)
            elif action == "clear":
                self.countdown_var.set("倒计时：--:--")
                if self.proc and self.proc.poll() is None and not self.is_paused:
                    self.countdown_reason_var.set("正在执行下一步")
            return True
        except Exception:
            return True

    def reset_countdown_display(self, status="状态：空闲", reason="等待任务开始"):
        self.is_paused = False
        self.status_var.set(status)
        self.countdown_var.set("倒计时：--:--")
        self.countdown_reason_var.set(reason)

    def reader(self, proc):
        try:
            for line in proc.stdout:
                self.q.put(line)
        except Exception as exc:
            self.q.put("读取日志失败：" + repr(exc) + "\n")

    def poll(self):
        while not self.q.empty():
            line = self.q.get()
            if not self.handle_worker_event(line):
                self.write_ui(line)
        if self.proc and self.proc.poll() is None:
            self.root.after(200, self.poll)
        elif self.proc:
            code = self.proc.returncode
            mode = self.proc_mode
            self.write_ui(f"\n进程已结束，退出码：{code}\n")
            self.proc = None
            self.proc_mode = ""
            if mode == "worker":
                if code == 0 and not self.telemetry_failed_accounts:
                    telemetry_status = "succeeded"
                elif code == 0 and self.telemetry_success_by_account:
                    telemetry_status = "partial"
                else:
                    telemetry_status = "failed"
                self.finish_task_telemetry(telemetry_status)
                self.reset_countdown_display("状态：任务完成" if code == 0 else "状态：任务异常", "全部浏览器已处理完成" if code == 0 else "请查看日志和调试截图")
                if code != 0:
                    try:
                        self.root.deiconify()
                        self.root.lift()
                    except Exception:
                        pass
                    messagebox.showwarning(
                        "任务已自动暂停",
                        f"发布任务进程异常退出（代码 {code}），软件已自动暂停。\n\n"
                        "请查看固定运行日志和调试截图。",
                        parent=self.root,
                    )
            else:
                self.reset_countdown_display("状态：空闲", "浏览器打开命令已完成")

    def run_subprocess(self, args, mode="helper"):
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"
        env["DOUYIN_GUI_PARENT"] = "1"
        self.proc = subprocess.Popen(
            args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1,
            cwd=str(APP_DIR), env=env, **hidden_subprocess_kwargs()
        )
        self.proc_mode = mode
        threading.Thread(target=self.reader, args=(self.proc,), daemon=True).start()
        self.root.after(200, self.poll)

    def open_browser_window(self):
        if self.proc and self.proc.poll() is None:
            messagebox.showwarning("正在运行", "请先等待当前命令结束或停止发布任务。")
            return
        if not self.save_from_ui(show_message=False):
            return
        index = self.selected_browser_index()
        account = self.browser_accounts[index if index is not None else 0]
        cmd = application_command(
            "--open-browser", CONFIG_PATH, "--account-id", account["id"]
        )
        self.write_ui(f"\n正在打开：{account['name']}（端口 {account['cdp_port']}）\n")
        self.reset_countdown_display(f"状态：打开 {account['name']}", "正在连接浏览器")
        self.run_subprocess(cmd, mode="open-browser")

    def probe_all_browsers(self):
        if self.proc and self.proc.poll() is None:
            messagebox.showwarning("正在运行", "请先等待当前命令结束或停止发布任务。")
            return
        if not self.save_from_ui(show_message=False):
            return
        for account in self.browser_accounts:
            self.browser_statuses[account["id"]] = "等待检查" if account.get("enabled", True) else "未启用"
        self.refresh_browser_tree()
        cmd = application_command("--probe-accounts", CONFIG_PATH)
        self.write_ui("\n开始检查已勾选的浏览器账号：只进入发布页，不上传、不填写、不发布。\n")
        self.reset_countdown_display("状态：检查账号", "准备第一个浏览器")
        self.run_subprocess(cmd, mode="probe")

    def start(self, scheduled=False):
        if self.proc and self.proc.poll() is None:
            if scheduled:
                self.write_ui("自动启动触发时已有任务运行，本次触发已跳过。\n")
            else:
                messagebox.showwarning("正在运行", "任务已经在运行。如需修改配置，请先停止任务。")
            return False
        if not self.auth_client.can_start_new_task:
            message = (
                "当前软件授权不允许开始新任务。请先登录有效的软件用户账号，"
                "并确认后台已为该账号开通 publish.enabled 发布权益。"
            )
            self.write_ui(message + "\n")
            if not scheduled:
                messagebox.showwarning("软件授权未就绪", message)
            return False
        if not self.save_from_ui(show_message=False):
            if not scheduled:
                messagebox.showwarning("无法开始", "当前配置无效，请查看日志。")
            return False
        try:
            for account in browser_accounts_from_config(self.cfg, enabled_only=True):
                account_cfg = config_for_browser_account(self.cfg, account)
                validate_publish_records_and_images(account_cfg)
                if not Path(account["browser_path"]).exists():
                    raise FileNotFoundError(f"浏览器账号“{account['name']}”路径不存在：{account['browser_path']}")
        except Exception as exc:
            self.write_ui(f"任务启动前检查失败：{exc}\n")
            if not scheduled:
                messagebox.showwarning("无法开始", str(exc))
            return False
        clear_pause_flags_for_queue(
            self.cfg, browser_accounts_from_config(self.cfg, enabled_only=True)
        )
        # 未点击“重置发布进度”时，开始/重启任务都继续显示并使用上一轮累计进度。
        self.browser_statuses = {}
        self.browser_success_counts = {}
        self.restore_browser_runtime_state()
        for account in self.browser_accounts:
            account_id = str(account["id"])
            if not account.get("enabled", True):
                self.browser_statuses[account_id] = "未启用"
            else:
                self.browser_statuses.setdefault(account_id, "等待任务")
                self.browser_success_counts.setdefault(account_id, 0)
        self.refresh_browser_tree()
        cmd = application_command("--worker", CONFIG_PATH)
        self.write_ui(("自动启动已触发。\n" if scheduled else "开始本轮发布。\n") + "已加载最新配置，浏览器将按列表顺序执行。\n")
        self.reset_countdown_display("状态：运行中", "准备第一个浏览器")
        self.reset_task_telemetry()
        self.run_subprocess(cmd, mode="worker")
        return True

    def retry_selected_failed_accounts(self):
        """重新执行已勾选且上次失败的账号；复用原队列轮次与累计进度。"""
        if self.proc and self.proc.poll() is None:
            messagebox.showwarning("任务运行中", "请先等待当前任务结束，或暂停/停止后再重试失败账号。")
            return False
        if not self.save_from_ui(show_message=False):
            return False

        queue_state = read_state(self.cfg)
        queue_run_id = str(queue_state.get("queue_run_id") or "").strip()
        failed_accounts = []
        for account in browser_accounts_from_config(self.cfg, enabled_only=True):
            account_cfg = config_for_browser_account(self.cfg, account)
            account_state = read_state(account_cfg)
            if queue_run_id and str(account_state.get("queue_run_id") or "") != queue_run_id:
                continue
            if str(account_state.get("last_account_outcome") or "") in {"error", "logged_out", "daily_limit"}:
                failed_accounts.append(account)

        if not failed_accounts:
            messagebox.showinfo("没有可重试账号", "当前勾选的账号中没有保存的失败、掉号或今日上限状态。")
            return False

        names = "、".join(str(account.get("name") or "未命名账号") for account in failed_accounts)
        self.write_ui(
            f"准备重试勾选的失败账号：{names}。原发布进度、当前图片和文案均继续保留。\n"
        )
        return self.start(scheduled=False)

    def toggle_pause(self):
        if not (self.proc and self.proc.poll() is None and self.proc_mode == "worker"):
            messagebox.showinfo("当前无任务", "当前没有正在运行的发布任务。")
            return
        # 暂停期间允许把全部账号取消勾选；继续后父调度器会读取空启用队列并安全结束本轮。
        was_paused = bool(self.is_paused or any(Path(flag).exists() for flag in self.active_pause_flags))
        if not self.save_from_ui(show_message=False, allow_no_enabled=was_paused):
            return
        current_flags = pause_flag_paths_for_queue(
            self.cfg, browser_accounts_from_config(self.cfg, enabled_only=True)
        )

        # 是否已暂停以 GUI 状态 + 已保存 flag 为准。队列在暂停期间可能被删除/改名/重排，
        # 不能只根据“当前队列”重新算 flag，否则旧账号 worker 的暂停文件可能永远残留。
        if was_paused:
            all_flags = []
            seen = set()
            for flag in list(self.active_pause_flags) + list(current_flags):
                path = Path(flag)
                key = os.path.normcase(os.path.abspath(str(path)))
                if key in seen:
                    continue
                seen.add(key)
                all_flags.append(path)
            for flag in all_flags:
                try:
                    if flag.exists():
                        flag.unlink()
                except OSError:
                    pass
            self.active_pause_flags = []
            self.is_paused = False
            self.status_var.set("状态：运行中")
            self.countdown_reason_var.set("等待工作进程继续；暂停期间的队列修改已保存")
            self.write_ui("已继续脚本；暂停期间的浏览器队列修改已保存，后续未启动队列项会读取最新配置。\n")
        else:
            created_flags = []
            for flag in current_flags:
                flag = Path(flag)
                flag.parent.mkdir(parents=True, exist_ok=True)
                flag.write_text("paused", encoding="utf-8")
                created_flags.append(flag)
            self.active_pause_flags = created_flags
            self.is_paused = True
            self.status_var.set("状态：已暂停")
            current = self.countdown_var.get()
            if current.startswith("倒计时："):
                self.countdown_var.set(current.replace("倒计时：", "暂停于：", 1))
            self.countdown_reason_var.set("暂停中：现在可以修改浏览器账号队列")
            self.write_ui("已请求暂停；当前安全步骤结束后生效。暂停期间可修改浏览器队列，点击继续时自动保存。\n")

    def stop(self):
        was_worker = bool(self.proc and self.proc.poll() is None and self.proc_mode == "worker")
        try:
            if self.proc and self.proc.poll() is None:
                if os.name == "nt":
                    subprocess.run(
                        ["taskkill", "/PID", str(self.proc.pid), "/T", "/F"],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        **hidden_subprocess_kwargs(),
                    )
                else:
                    self.proc.terminate()
                try:
                    self.proc.wait(timeout=5)
                except Exception:
                    self.proc.kill()
                    self.proc.wait(timeout=3)
                self.write_ui("已停止当前任务进程。\n")
            else:
                self.write_ui("当前没有正在运行的任务。\n")
        except Exception as exc:
            self.write_ui("停止任务失败：" + repr(exc) + "\n")
        finally:
            self.proc = None
            self.proc_mode = ""
            clear_pause_flags_for_queue(self.cfg, self.browser_accounts)
            for flag in list(self.active_pause_flags):
                try:
                    path = Path(flag)
                    if path.exists():
                        path.unlink()
                except OSError:
                    pass
            self.active_pause_flags = []
            if was_worker:
                self.finish_task_telemetry("stopped")
            self.reset_countdown_display("状态：已停止", "任务已停止")

    def reset_state(self):
        if not self.save_from_ui(show_message=False):
            return
        accounts = browser_accounts_from_config(self.cfg)
        state_configs = [self.cfg] + [config_for_browser_account(self.cfg, account) for account in accounts]
        unique_configs = {}
        for state_cfg in state_configs:
            unique_configs[str(Path(state_cfg["state_path"]).resolve())] = state_cfg
        if messagebox.askyesno(
            "确认重置",
            f"确定重置全部 {len(accounts)} 个队列项的文案和独立排期进度吗？\n"
            "不会删除图片、Excel、日志或调试截图。",
        ):
            for state_cfg in unique_configs.values():
                state = read_state(state_cfg)
                state.update({
                    "copy_index": 0,
                    "slot_index": 0,
                    "used_copy_keys": [],
                    "schedule_signature": "",
                    "queue_run_id": "",
                    "queue_run_active": False,
                    "queue_run_signature": "",
                    "queue_completed_account_ids": [],
                    "run_progress": 0,
                    "run_quota": 0,
                    "run_active": False,
                })
                write_state(state_cfg, state)
            clear_pause_flags_for_queue(self.cfg, accounts)
            self.write_ui(f"已重置 {len(accounts)} 个队列项的独立发布进度。\n")

    def poll_weekly_scheduler(self):
        try:
            cfg = self.collect_cfg()
            next_trigger = next_weekly_trigger(cfg)
            next_text = (
                f"{next_trigger:%Y-%m-%d %H:%M} {WEEKDAY_LABELS[next_trigger.weekday()]}"
                if next_trigger else "未启用"
            )
            self.next_schedule_var.set(
                "下次自动启动：" + next_text
            )
            key = weekly_trigger_key(cfg)
            if key:
                state = read_state(cfg)
                if state.get("weekly_last_trigger") != key:
                    state["weekly_last_trigger"] = key
                    write_state(cfg, state)
                    if self.proc and self.proc.poll() is None:
                        self.write_ui(f"[{key}] 自动启动触发，但已有任务运行，本次已跳过。\n")
                    else:
                        self.write_ui(f"[{key}] 到达自动启动时间。\n")
                        self.start(scheduled=True)
        except Exception as exc:
            self.next_schedule_var.set(f"自动启动配置待修正：{exc}")
        finally:
            try:
                self.root.after(5000, self.poll_weekly_scheduler)
            except Exception:
                pass

    def open_log_folder(self):
        try:
            cfg = self.collect_cfg()
            path = Path(cfg.get("log_path", "")).parent
            path.mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))
        except Exception as exc:
            messagebox.showerror("无法打开日志目录", str(exc))

    def on_close(self):
        if self.proc and self.proc.poll() is None:
            if not messagebox.askyesno("退出程序", "发布任务仍在运行，退出会停止任务。确定退出吗？"):
                return
            self.stop()
        self.persist_browser_tree_column_widths()
        self.root.destroy()
def run_gui():
    hide_own_console_window()
    if tk is None:
        print("当前 Python 不支持 tkinter，请安装标准版 Python。")
        return
    root = tk.Tk()
    App(root)
    root.mainloop()


def run_self_test():
    """不连接浏览器、不发布内容的内置冒烟测试。"""
    assert APP_VERSION == "3.0.4"
    assert APP_NAME.endswith(APP_VERSION)
    empty_release = platform_release_to_manifest({"version": {}})
    assert empty_release["latest_version"] == ""
    valid_release = platform_release_to_manifest(
        {
            "version": {
                "latestVersion": "2.5.3",
                "downloadUri": "https://example.invalid/update.zip",
                "fileSha256": "a" * 64,
            }
        }
    )
    assert valid_release["latest_version"] == "2.5.3"
    assert valid_release["sha256"] == "a" * 64
    body, topics = split_text_topics("正文内容 ##话题一 #话题二")
    assert "#" not in body and topics == ["话题一", "话题二"]
    assert COPY_FILE_EXTS == {".xlsx", ".xls", ".xlsm"}
    assert extract_copy_texts(pd.DataFrame([["第一条文案"], ["第二条文案"]])) == ["第一条文案", "第二条文案"]
    assert extract_copy_texts(pd.DataFrame([["第一条文案", "备注一"], ["第二条文案", "备注二"]])) == ["第一条文案", "第二条文案"]
    assert extract_copy_texts(pd.DataFrame([["文案"], ["第一条文案"], ["第二条文案"]])) == ["第一条文案", "第二条文案"]
    structured_records = extract_publish_records(pd.DataFrame([
        ["城市", "标题", "文案"],
        ["北京", "北京标题", "北京正文 #北京"],
        ["上海", "", "上海正文"],
    ]))
    assert structured_records[0]["city"] == "北京"
    assert structured_records[0]["title"] == "北京标题"
    assert structured_records[0]["copy"] == "北京正文 #北京"
    assert structured_records[0]["excel_row"] == 2
    assert structured_records[0]["has_city_column"] is True
    assert structured_records[1]["title"] == ""
    alias_records = extract_publish_records(pd.DataFrame([
        ["城市名称", "标题内容", "内容文案"],
        ["北京市", "北京标题", "北京正文"],
    ]))
    assert alias_records[0]["city"] == "北京市"
    assert alias_records[0]["title"] == "北京标题"
    assert alias_records[0]["copy"] == "北京正文"
    headerless_city_records = extract_publish_records(
        pd.DataFrame([
            ["济南", "标题一", "正文一"],
            ["青岛", "标题二", "正文二"],
        ]),
        known_city_names=["济南", "青岛"],
    )
    assert headerless_city_records[0]["city"] == "济南"
    assert headerless_city_records[0]["title"] == "标题一"
    assert headerless_city_records[0]["copy"] == "正文一"
    assert headerless_city_records[0]["format"] == "headerless_city"
    headerless_city_copy_records = extract_publish_records(
        pd.DataFrame([["济南", "正文一"], ["青岛", "正文二"]]),
        known_city_names=["济南", "青岛"],
    )
    assert headerless_city_copy_records[0]["title"] == ""
    assert headerless_city_copy_records[0]["copy"] == "正文一"
    assert extract_copy_texts(pd.DataFrame([["标题"], ["旧版第一条"], ["旧版第二条"]])) == ["旧版第一条", "旧版第二条"]
    try:
        extract_publish_records(pd.DataFrame([["城市", "标题"], ["北京", "标题一"]]))
        raise AssertionError("有城市/标题表头时必须包含文案列")
    except RuntimeError:
        pass
    with tempfile.TemporaryDirectory(prefix="douyin_copy_selftest_") as temp_dir:
        image_root = Path(temp_dir) / "images"
        city_dir = image_root / "北京"
        city_dir.mkdir(parents=True)
        (city_dir / "b.jpg").write_bytes(b"b")
        (city_dir / "A.png").write_bytes(b"a")
        image_cfg = {"image_dir": str(image_root), "random_image": False}
        assert choose_image_for_record(image_cfg, structured_records[0]).name == "A.png"
        assert choose_image_for_record({**image_cfg, "random_image": True}, structured_records[0]).parent == city_dir
        assert resolve_image_dir_for_record(image_cfg, alias_records[0]) == city_dir
        try:
            list_images_for_record(image_cfg, {**structured_records[0], "city": "../北京"})
            raise AssertionError("城市列不应允许路径跳转")
        except RuntimeError:
            pass
        try:
            list_images_for_record(image_cfg, {**structured_records[1], "city": ""})
            raise AssertionError("存在城市列时城市不能为空")
        except RuntimeError:
            pass

        headerless_path = Path(temp_dir) / "headerless.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "sheet1"
        sheet.append(["第一条文案"])
        sheet.append(["第二条文案"])
        workbook.save(headerless_path)
        workbook.close()
        copy_cfg = {"excel_path": str(headerless_path), "sheet_name": "sheet1"}
        assert read_copies(copy_cfg) == ["第一条文案", "第二条文案"]
        assert delete_copy_from_excel(copy_cfg, "第一条文案") is True
        assert read_copies(copy_cfg) == ["第二条文案"]

        headerless_city_path = Path(temp_dir) / "headerless_city.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "sheet1"
        sheet.append(["北京", "标题一", "正文一"])
        sheet.append(["北京市", "", "正文二"])
        workbook.save(headerless_city_path)
        workbook.close()
        headerless_city_cfg = {
            "excel_path": str(headerless_city_path), "sheet_name": "sheet1",
            "image_dir": str(image_root), "random_image": False,
        }
        loaded_headerless_city = validate_publish_records_and_images(headerless_city_cfg)
        assert len(loaded_headerless_city) == 2
        assert loaded_headerless_city[0]["city"] == "北京"
        assert loaded_headerless_city[0]["title"] == "标题一"
        assert loaded_headerless_city[1]["city"] == "北京市"

        header_path = Path(temp_dir) / "with_header.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "sheet1"
        sheet.append(["文案"])
        sheet.append(["第一条文案"])
        workbook.save(header_path)
        workbook.close()
        assert read_copies({"excel_path": str(header_path), "sheet_name": "sheet1"}) == ["第一条文案"]

        structured_path = Path(temp_dir) / "city_title_copy.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "sheet1"
        sheet.append(["城市", "标题", "文案"])
        sheet.append(["北京", "标题一", "相同正文"])
        sheet.append(["北京", "标题二", "相同正文"])
        workbook.save(structured_path)
        workbook.close()
        structured_cfg = {
            "excel_path": str(structured_path), "sheet_name": "sheet1",
            "image_dir": str(image_root), "random_image": False,
        }
        records = read_publish_records(structured_cfg)
        assert len(records) == 2 and records[0]["title"] == "标题一"
        assert publish_record_key(records[0]) != publish_record_key(records[1])
        assert filter_used_publish_records(records, {"used_copy_keys": [publish_record_key(records[0])]}) == [records[1]]
        assert delete_copy_from_excel(structured_cfg, records[0]) is True
        remaining = read_publish_records(structured_cfg)
        assert len(remaining) == 1 and remaining[0]["title"] == "标题二"
    legacy = default_config()
    legacy["browser_path"] = r"D:\浏览器7.lnk"
    legacy["image_dir"] = r"D:\共享城市图片"
    legacy["browser_account_suggestions"] = []
    accounts = browser_accounts_from_config(legacy)
    assert len(accounts) == 1 and accounts[0]["posts_per_run"] == 1
    custom = normalize_browser_account(
        {
            "id": "browser_custom",
            "name": "账号专属内容",
            "browser_path": r"D:\浏览器9.lnk",
            "independent_content": True,
            "image_dir": r"D:\账号9图片",
            "excel_path": r"D:\账号9文案.xlsx",
            "sheet_name": "账号9",
        },
        0,
        legacy,
    )
    mapped = config_for_browser_account(legacy, custom)
    assert "image_dir" not in custom
    assert mapped["image_dir"] == r"D:\共享城市图片"
    assert mapped["excel_path"] == r"D:\账号9文案.xlsx"
    assert "browser_custom" in mapped["state_path"]
    duplicate_a = normalize_browser_account(
        {**custom, "id": "queue_a", "name": "账号第1次", "cdp_port": 9229}, 0, legacy
    )
    duplicate_b = normalize_browser_account(
        {**custom, "id": "queue_b", "name": "账号第2次", "cdp_port": 9229}, 1, legacy
    )
    validate_browser_queue_ports([duplicate_a, duplicate_b])
    duplicate_cfg = {**legacy, "browser_accounts": [duplicate_a, {**duplicate_b, "id": "queue_a"}]}
    duplicate_accounts = browser_accounts_from_config(duplicate_cfg)
    assert len({item["id"] for item in duplicate_accounts}) == 2
    duplicate_state_a = config_for_browser_account(legacy, duplicate_accounts[0])
    duplicate_state_b = config_for_browser_account(legacy, duplicate_accounts[1])
    assert duplicate_state_a["state_path"] != duplicate_state_b["state_path"]
    assert config_for_browser_account(duplicate_state_a, duplicate_accounts[0])["state_path"] == duplicate_state_a["state_path"]
    try:
        validate_browser_queue_ports([duplicate_a, {**duplicate_b, "browser_path": r"D:\浏览器8.lnk"}])
        raise AssertionError("不同浏览器不应允许复用同一 CDP 端口")
    except RuntimeError:
        pass
    independent_schedule = normalize_browser_account(
        {
            **duplicate_a,
            "publish_start_date": "2026-08-05",
            "publish_end_date": "2026-08-05",
            "start_hour": 8,
            "end_hour": 8,
            "per_hour_count": 2,
            "custom_minutes": "5,35",
            "posts_per_run": 2,
        },
        0,
        legacy,
    )
    independent_cfg = config_for_browser_account(legacy, independent_schedule)
    assert len(build_slots(independent_cfg, log_result=False)) == 2
    assert independent_cfg["publish_start_date"] == "2026-08-05"
    repeat_slots = build_slots(independent_cfg, log_result=False)
    first_slot, first_cycle, first_offset = repeating_schedule_slot(repeat_slots, 0)
    repeated_slot, repeated_cycle, repeated_offset = repeating_schedule_slot(repeat_slots, len(repeat_slots))
    overflow_slot, overflow_cycle, overflow_offset = repeating_schedule_slot(repeat_slots, 5)
    assert first_slot == repeated_slot and (first_cycle, first_offset) == (1, 0)
    assert (repeated_cycle, repeated_offset) == (2, 0)
    assert overflow_slot == repeat_slots[1] and (overflow_cycle, overflow_offset) == (3, 1)
    assert account_run_resume_progress({"slot_index": 64}, "queue_test", 75) == 64
    assert account_run_resume_progress({"slot_index": 75}, "queue_test", 75) == 0
    assert account_run_resume_progress(
        {"queue_run_id": "queue_test", "run_progress": 64}, "queue_test", 75
    ) == 64
    assert account_run_resume_progress(
        {"queue_run_id": "queue_old", "run_progress": 64}, "queue_new", 75
    ) == 0
    assert normalize_rotate_batch_size(3) == 3
    assert normalize_rotate_batch_size(0) == 1
    assert normalize_rotate_batch_size("bad") == 1
    assert rotation_batch_target(12, 0, 5) == 5
    assert rotation_batch_target(12, 5, 5) == 5
    assert rotation_batch_target(12, 10, 5) == 2
    assert normalize_daily_limit_switch_wait_seconds("15") == 15
    assert normalize_daily_limit_switch_wait_seconds("-3") == 0
    assert text_indicates_daily_publish_limit("抱歉，今天投稿次数已达到上限，请明天再试")
    assert text_indicates_daily_publish_limit("抱歉，今天投稿次数已达到上限\n请明天再试")
    retry_cfg = {
        "retry_times": 2,
        "overall_retry_times": 1,
        "wait_min_seconds": 0.01,
        "wait_max_seconds": 0.01,
    }
    retry_calls = {"step": 0, "overall": 0}

    def eventually_succeeds():
        retry_calls["step"] += 1
        if retry_calls["step"] < 3:
            raise RuntimeError("示例步骤暂未完成")
        return "ok"

    assert run_publish_workflow(
        retry_cfg,
        lambda _attempt, _total: run_publish_step(retry_cfg, None, "示例步骤", eventually_succeeds),
        lambda _attempt, _error: retry_calls.__setitem__("overall", retry_calls["overall"] + 1),
    ) == "ok"
    assert retry_calls == {"step": 3, "overall": 0}

    exhausted_calls = {"step": 0, "overall": 0}

    def always_fails():
        exhausted_calls["step"] += 1
        raise RuntimeError("持续失败根因")

    try:
        run_publish_workflow(
            retry_cfg,
            lambda _attempt, _total: run_publish_step(retry_cfg, None, "音乐选择", always_fails),
            lambda _attempt, _error: exhausted_calls.__setitem__("overall", exhausted_calls["overall"] + 1),
        )
        raise AssertionError("两级重试耗尽后必须抛出 PublishWorkflowFailedError")
    except PublishWorkflowFailedError as exc:
        assert exhausted_calls == {"step": 6, "overall": 1}
        assert exc.step_name == "音乐选择"
        assert "持续失败根因" in concise_publish_failure(exc)

    daily_limit_calls = {"count": 0}

    def raises_daily_limit():
        daily_limit_calls["count"] += 1
        raise DailyPublishLimitError("今天投稿次数已达到上限")

    try:
        run_publish_workflow(
            retry_cfg,
            lambda _attempt, _total: run_publish_step(retry_cfg, None, "提交发布", raises_daily_limit),
        )
        raise AssertionError("今日投稿上限必须直接向上抛出")
    except DailyPublishLimitError:
        assert daily_limit_calls["count"] == 1

    with tempfile.TemporaryDirectory(prefix="douyin_debug_cleanup_selftest_") as temp_dir:
        debug_dir = Path(temp_dir) / "screenshots"
        debug_dir.mkdir()
        generated = debug_dir / "20260810_120000_step_failed.png"
        unrelated = debug_dir / "keep.png"
        nested_dir = debug_dir / "nested"
        nested_dir.mkdir()
        nested_generated = nested_dir / "20260810_120001_nested.png"
        generated.write_bytes(b"generated")
        unrelated.write_bytes(b"unrelated")
        nested_generated.write_bytes(b"nested")
        cleanup_state = Path(temp_dir) / "cleanup_state.json"
        cleanup_cfg = {"debug_dir": str(debug_dir)}
        assert maintain_debug_screenshots(cleanup_cfg, cleanup_state)["cleaned"] is False
        assert maintain_debug_screenshots(cleanup_cfg, cleanup_state)["cleaned"] is False
        cleanup_result = maintain_debug_screenshots(cleanup_cfg, cleanup_state)
        assert cleanup_result["cleaned"] is True and cleanup_result["removed"] == 1
        assert not generated.exists() and unrelated.exists() and nested_generated.exists()
    assert rotation_batch_target(12, 12, 5) == 0
    assert normalize_browser_queue_column_widths({"path": 333})["path"] == 333
    assert normalize_browser_queue_column_widths({"path": 1})["path"] == BROWSER_QUEUE_MIN_WIDTHS["path"]
    with tempfile.TemporaryDirectory(prefix="douyin_resume_selftest_") as temp_dir:
        queue_cfg = {"state_path": str(Path(temp_dir) / "queue_state.json")}
        queue_accounts = [duplicate_a, duplicate_b]
        queue_run_id, completed_ids, resumed = prepare_queue_run(queue_cfg, queue_accounts)
        assert resumed is False and completed_ids == set()
        update_queue_run_state(
            queue_cfg, queue_run_id, queue_accounts, {duplicate_a["id"]}, active=True
        )
        resumed_id, resumed_completed, resumed = prepare_queue_run(queue_cfg, queue_accounts)
        assert resumed is True and resumed_id == queue_run_id
        assert resumed_completed == {duplicate_a["id"]}
        pause_paths = pause_flag_paths_for_queue(queue_cfg, queue_accounts)
        for flag in pause_paths:
            flag.parent.mkdir(parents=True, exist_ok=True)
            flag.write_text("paused", encoding="utf-8")
        clear_pause_flags_for_queue(queue_cfg, queue_accounts)
        assert all(not flag.exists() for flag in pause_paths)
        update_queue_run_state(
            queue_cfg, queue_run_id, queue_accounts,
            {duplicate_a["id"], duplicate_b["id"]}, active=True,
        )
        queue_cfg["browser_accounts"] = queue_accounts
        queue_config_path = Path(temp_dir) / "queue_config.json"
        queue_config_path.write_text(
            json.dumps(queue_cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        run_isolated_browser_sequence(queue_config_path, "--unused", publish_mode=True)
        assert read_state(queue_cfg)["queue_run_active"] is False
        next_run_id, next_completed, resumed = prepare_queue_run(queue_cfg, queue_accounts)
        assert resumed is False and next_run_id != queue_run_id and next_completed == set()
    visible_cmd = build_browser_launch_command("chrome.exe", "--foo", 9222, "https://example.com", headless=False)
    hidden_cmd = build_browser_launch_command("chrome.exe", "--foo", 9222, "https://example.com", headless=True)
    assert "--headless" not in visible_cmd
    assert "--headless=new" in hidden_cmd and "--window-size=1440,1100" in hidden_cmd

    schedule_cfg = {
        "weekly_start_enabled": True,
        "weekly_start_days": [0],
        "weekly_start_time": "09:30",
    }
    monday = datetime(2026, 8, 3, 9, 30, 20)
    assert weekly_trigger_key(schedule_cfg, monday) == "2026-08-03 09:30"
    assert weekly_trigger_key(schedule_cfg, monday + timedelta(minutes=1)) == ""
    next_run = next_weekly_trigger(schedule_cfg, datetime(2026, 8, 3, 9, 29, 59))
    assert next_run == datetime(2026, 8, 3, 9, 30)
    schedule_slot = datetime(2026, 8, 11, 8, 5)
    assert normalize_schedule_input_value("2026/8/11 8:05:00") == "2026-08-11 08:05"
    assert schedule_value_matches("2026-08-11T08:05:00", schedule_slot)
    assert not schedule_value_matches("2026-08-11 17:05", schedule_slot)

    class _AlreadyAtBottomMouse:
        def __init__(self):
            self.wheel_calls = 0

        def wheel(self, _x, _y):
            self.wheel_calls += 1

    class _AlreadyAtBottomPage:
        def __init__(self):
            self.mouse = _AlreadyAtBottomMouse()
            self.wait_calls = 0

        def evaluate(self, _script, *_args):
            return {
                "publishSettingVisible": True,
                "publishTimeVisible": True,
                "publishButtonVisible": True,
                "scrollY": 1000,
                "innerHeight": 700,
                "bodyHeight": 1700,
            }

        def wait_for_timeout(self, _milliseconds):
            self.wait_calls += 1

    bottom_page = _AlreadyAtBottomPage()
    assert scroll_to_publish_settings({"step_wait_min": 0, "step_wait_max": 0}, bottom_page)
    assert bottom_page.mouse.wheel_calls == 0 and bottom_page.wait_calls == 0
    if tk is not None:
        root = tk.Tk()
        root.withdraw()
        gui = App(root)
        root.update_idletasks()
        assert root.state() == "withdrawn"
        assert gui.main_window_revealed is False
        assert len(gui.tab_scroll_areas) == 4
        assert len({id(canvas) for canvas in gui.tab_scroll_canvases}) == 4
        assert not root.bind_all("<MouseWheel>")
        assert gui.start_button.master is gui.task_controls_frame
        assert gui.retry_failed_button.master is gui.task_controls_frame
        assert gui.stop_button.master is gui.task_controls_frame
        assert gui.reset_progress_button.master is not gui.task_controls_frame
        assert gui.clear_log_button.master is gui.reset_progress_button.master
        assert [gui.notebook.tab(index, "text") for index in range(gui.notebook.index("end"))] == [
            "发布中心", "账号与排期", "自动启动", "系统与日志"
        ]
        assert "retry_times" in gui.vars
        assert "overall_retry_times" in gui.vars
        assert "image_dir" in gui.vars
        assert not {"excel_path", "publish_start_date", "publish_end_date", "max_posts_this_run"} & set(gui.vars)
        assert "use_schedule" not in gui.bool_vars
        assert "rotate_accounts_each_post" in gui.bool_vars
        assert "rotate_batch_size" in gui.vars
        expected_rotate_state = "normal" if gui.bool_vars["rotate_accounts_each_post"].get() else "disabled"
        assert str(gui.rotate_batch_spinbox.cget("state")) == expected_rotate_state
        gui.bool_vars["rotate_accounts_each_post"].set(False)
        root.update_idletasks()
        assert str(gui.rotate_batch_spinbox.cget("state")) == "disabled"
        assert int(gui.browser_tree.cget("height")) == 8
        assert tuple(gui.browser_tree.cget("columns"))[0] == "publish"
        assert "published" in tuple(gui.browser_tree.cget("columns"))
        assert "image" not in tuple(gui.browser_tree.cget("columns"))
        assert all(not bool(gui.browser_tree.column(key, "stretch")) for key in BROWSER_QUEUE_COLUMNS)
        saved_width_configs = []
        old_save_config = globals()["save_config"]
        try:
            globals()["save_config"] = lambda cfg: saved_width_configs.append(copy.deepcopy(cfg))
            gui.browser_tree.column("path", width=333)
            assert gui.persist_browser_tree_column_widths() is True
            assert saved_width_configs[-1]["browser_queue_column_widths"]["path"] == 333
            assert gui.persist_browser_tree_column_widths() is False
        finally:
            globals()["save_config"] = old_save_config
        first_account_id = gui.browser_accounts[0]["id"]
        gui.handle_worker_event(
            BROWSER_STATUS_EVENT_PREFIX + json.dumps({
                "id": first_account_id,
                "name": gui.browser_accounts[0]["name"],
                "status": "error",
                "message": "音乐选择失败：未检测到已使用音乐；进度 2/5 已保留",
                "success_count": 2,
                "quota": 5,
            }, ensure_ascii=False)
        )
        assert "音乐选择失败" in gui.browser_statuses[first_account_id]
        assert gui.browser_success_counts[first_account_id] >= 2
        first_enabled = parse_bool(gui.browser_accounts[0].get("enabled"), True)
        assert gui.toggle_browser_account_enabled(first_account_id) is True
        assert parse_bool(gui.browser_accounts[0].get("enabled"), True) is not first_enabled
        assert gui.set_all_browser_accounts_enabled(True) is True
        assert len(browser_accounts_from_config({**gui.cfg, "browser_accounts": gui.browser_accounts}, enabled_only=True)) == len(gui.browser_accounts)
        assert gui.set_all_browser_accounts_enabled(False) is True
        assert browser_accounts_from_config({**gui.cfg, "browser_accounts": gui.browser_accounts}, enabled_only=True) == []
        disabled_cfg = gui.collect_cfg(allow_no_enabled=True)
        assert browser_accounts_from_config(disabled_cfg, enabled_only=True) == []
        gui.set_all_browser_accounts_enabled(True)
        with tempfile.TemporaryDirectory(prefix="douyin_queue_selftest_") as temp_dir:
            queue_path = Path(temp_dir) / "queue.xlsx"
            old_save_dialog = filedialog.asksaveasfilename
            old_open_dialog = filedialog.askopenfilename
            old_showinfo = messagebox.showinfo
            old_showerror = messagebox.showerror
            old_askyesno = messagebox.askyesno
            dialog_errors = []
            try:
                gui.browser_accounts = [duplicate_a, duplicate_b]
                gui.refresh_browser_tree()
                filedialog.asksaveasfilename = lambda **_kwargs: str(queue_path)
                filedialog.askopenfilename = lambda **_kwargs: str(queue_path)
                messagebox.showinfo = lambda *_args, **_kwargs: None
                messagebox.showerror = lambda title, message, **_kwargs: dialog_errors.append((title, message))
                messagebox.askyesno = lambda *_args, **_kwargs: True
                gui.export_browser_queue_table()
                assert queue_path.is_file()
                gui.browser_accounts = []
                gui.refresh_browser_tree()
                gui.import_browser_queue_table()
                assert len(gui.browser_accounts) == 2
                assert gui.browser_accounts[0]["browser_path"] == gui.browser_accounts[1]["browser_path"]
                assert len({item["id"] for item in gui.browser_accounts}) == 2
                assert not dialog_errors
            finally:
                filedialog.asksaveasfilename = old_save_dialog
                filedialog.askopenfilename = old_open_dialog
                messagebox.showinfo = old_showinfo
                messagebox.showerror = old_showerror
                messagebox.askyesno = old_askyesno
        root.destroy()
    print("SELF_TEST_OK", flush=True)


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        run_self_test()
    elif "--account-worker" in sys.argv:
        account_worker(sys.argv[sys.argv.index("--account-worker") + 1])
    elif "--worker" in sys.argv:
        worker(sys.argv[sys.argv.index("--worker") + 1])
    elif "--probe-account" in sys.argv:
        probe_account_in_process(sys.argv[sys.argv.index("--probe-account") + 1])
    elif "--probe-accounts" in sys.argv:
        probe_browser_accounts(sys.argv[sys.argv.index("--probe-accounts") + 1])
    elif "--open-browser" in sys.argv:
        cfg = json.loads(Path(sys.argv[sys.argv.index("--open-browser") + 1]).read_text(encoding="utf-8"))
        if "--account-id" in sys.argv:
            account_id = sys.argv[sys.argv.index("--account-id") + 1]
            account = next((item for item in browser_accounts_from_config(cfg) if item.get("id") == account_id), None)
            if account is None:
                raise RuntimeError(f"找不到浏览器账号：{account_id}")
            cfg = config_for_browser_account(cfg, account)
        open_browser_with_cdp(cfg, force_new=False)
        wlog("浏览器前端窗口已打开。")
    else:
        run_gui()
