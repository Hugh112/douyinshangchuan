# -*- coding: utf-8 -*-
"""Simple online updater for Douyin Publisher."""
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
import zipfile
from pathlib import Path


OBSOLETE_PATHS = [
    "#U542f#U52a8#U6296#U97f3#U5e03#U5668_#U65e0#U9ed1#U6846.vbs",
    "BLACK_WINDOW_FIX_NOTES.txt",
    "README_v2.2.5.txt",
    "README_v2.2.6.txt",
    "README_v2.2.7.txt",
    "README_v2.2.8.txt",
    "README_v2.3.0.txt",
    "app_main.py",
    "app_main.pyw",
    "Start_Douyin_Publisher_v2.1.bat",
    "Start_Douyin_Publisher_v2.1.vbs",
    "Start_Douyin_Publisher_v2.2.bat",
    "Start_Douyin_Publisher_v2.2.vbs",
    "Start_Douyin_Publisher_v2.2.5.bat",
    "Start_Douyin_Publisher_v2.2.5.vbs",
    "Start_Douyin_Publisher_v2.2.6.bat",
    "Start_Douyin_Publisher_v2.2.6.vbs",
    "Start_Douyin_Publisher_v2.2.7.bat",
    "Start_Douyin_Publisher_v2.2.7.vbs",
    "debug_start_v2.1.bat",
    "debug_start_v2.2.bat",
    "debug_start_v2.2.5.bat",
    "debug_start_v2.2.6.bat",
    "debug_start_v2.2.7.bat",
    "使用说明.txt",
    "创建快捷方式.bat",
    "启动_抖音单图文发布v2.2.vbs",
    "安装依赖_可选.bat",
    "抖音单图文发布v2.2.lnk",
    "浏览器目录查找.txt",
    "在线更新说明.txt",
    "app/__pycache__",
    "app/app_main.spec",
    "app/抖音单图文发布v2.2.py",
    "app/抖音单图文发布v2.2.pyw",
    "app/本次修复说明.txt",
    "app/鎶栭煶鍗曞浘鏂囧彂甯僾2.1.py",
    "app/鎶栭煶鍗曞浘鏂囧彂甯僾2.1.pyw",
    "app/鏈淇璇存槑.txt",
    # v2.5.0 起以原生 EXE 启动；以下仅清理确定属于旧 Python/VBS 启动链的文件。
    "Start_Douyin_Publisher.bat",
    "start_app.bat",
    "start_app.vbs",
    "debug_start.bat",
    "requirements.txt",
    "app/start_app.bat",
    "app/start_app.vbs",
    "app/debug_start.bat",
    "app/app_main.py",
    "app/app_main.pyw",
    "app/online_updater.py",
    "app/assets/app_logo.png",
    "app/assets/app_logo.ico",
]


def hidden_subprocess_kwargs():
    """Prevent updater/relauncher console windows from appearing on Windows."""
    if os.name != "nt":
        return {}
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = getattr(subprocess, "SW_HIDE", 0)
    return {
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
        "startupinfo": startupinfo,
    }


def launch_python_executable(script_path):
    exe = Path(sys.executable)
    if os.name == "nt" and Path(script_path).suffix.lower() == ".pyw" and exe.name.lower() == "python.exe":
        pythonw = exe.with_name("pythonw.exe")
        if pythonw.exists():
            return str(pythonw)
    return str(exe)


def download_with_retry(url, dst, retry=3, timeout=60):
    parsed = urllib.parse.urlparse(str(url or ""))
    if parsed.scheme.lower() != "https":
        raise RuntimeError("更新包下载地址必须使用 HTTPS。")
    headers = {"User-Agent": "DouyinPublisherUpdater/2.5.1", "Cache-Control": "no-cache", "Pragma": "no-cache", "Connection": "close"}
    last_err = None
    sep = "&" if "?" in url else "?"
    url2 = f"{url}{sep}_t={int(time.time())}"
    for i in range(max(1, retry)):
        try:
            req = urllib.request.Request(url2, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as r, open(dst, "wb") as f:
                shutil.copyfileobj(r, f)
            return
        except Exception as e:
            last_err = e
            time.sleep(2 + i * 2)
    raise last_err


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def wait_process_exit(pid, seconds=8):
    try:
        pid = int(pid or 0)
    except Exception:
        pid = 0
    if not pid:
        return
    if os.name != "nt":
        time.sleep(2)
        return
    try:
        import ctypes
        SYNCHRONIZE = 0x00100000
        handle = ctypes.windll.kernel32.OpenProcess(SYNCHRONIZE, False, pid)
        if handle:
            ctypes.windll.kernel32.WaitForSingleObject(handle, int(seconds * 1000))
            ctypes.windll.kernel32.CloseHandle(handle)
        else:
            time.sleep(2)
    except Exception:
        time.sleep(2)


def copy_tree(src, dst, preserve_paths):
    src = Path(src)
    dst = Path(dst)
    preserve_paths = set(str(p).replace(chr(92), "/") for p in preserve_paths or [])
    backups = {}
    for rel in preserve_paths:
        p = dst / rel
        if p.exists():
            backups[rel] = p.read_bytes()
    for root, dirs, files in os.walk(src):
        rootp = Path(root)
        rel_root = rootp.relative_to(src)
        for name in files:
            src_file = rootp / name
            rel = (rel_root / name).as_posix()
            if rel in preserve_paths:
                continue
            dst_file = dst / rel
            dst_file.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src_file, dst_file)
    for rel, data in backups.items():
        p = dst / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data)


def validate_zip_members(archive):
    """更新包只允许相对普通文件，拒绝绝对路径、父目录穿越和符号链接。"""
    for info in archive.infolist():
        name = str(info.filename or "").replace("\\", "/")
        if not name or name.endswith("/"):
            continue
        parts = [part for part in name.split("/") if part not in {"", "."}]
        if name.startswith("/") or (parts and ":" in parts[0]) or ".." in parts:
            raise RuntimeError(f"更新包包含不安全路径：{info.filename}")
        unix_mode = (info.external_attr >> 16) & 0xF000
        if unix_mode == 0xA000:
            raise RuntimeError(f"更新包不允许符号链接：{info.filename}")


def remove_obsolete_paths(install_dir):
    root = Path(install_dir).resolve()
    for rel in OBSOLETE_PATHS:
        target = (root / rel).resolve()
        try:
            target.relative_to(root)
        except ValueError as exc:
            raise RuntimeError(f"无效的历史清理路径：{rel}") from exc
        if target.is_dir():
            shutil.rmtree(target, ignore_errors=True)
        elif target.exists():
            target.unlink()


def relaunch_without_console(install_dir, fallback_launch=""):
    fixed_exe = Path(install_dir) / "DouyinPublisher.exe"
    if os.name == "nt" and fixed_exe.exists():
        subprocess.Popen(
            [str(fixed_exe)],
            cwd=str(fixed_exe.parent),
            **hidden_subprocess_kwargs(),
        )
        return
    fixed_launcher = Path(install_dir) / "Start_Douyin_Publisher.vbs"
    if os.name == "nt" and fixed_launcher.exists():
        subprocess.Popen(
            ["wscript.exe", "//nologo", str(fixed_launcher)],
            cwd=str(fixed_launcher.parent),
            **hidden_subprocess_kwargs(),
        )
        return
    launch = Path(fallback_launch) if fallback_launch else Path(install_dir) / "app" / "app_main.pyw"
    if launch.name.lower() in {"app_main.py", "app_main.pyw"}:
        launch = Path(install_dir) / "app" / launch.name
    if launch.exists():
        subprocess.Popen(
            [launch_python_executable(launch), str(launch)],
            cwd=str(launch.parent),
            **hidden_subprocess_kwargs(),
        )


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--install-dir", required=True)
    ap.add_argument("--current-pid", default="0")
    ap.add_argument("--download-url", required=True)
    ap.add_argument("--sha256", required=True)
    ap.add_argument("--launch", default="")
    ap.add_argument("--preserve", default="[]")
    args = ap.parse_args()
    install_dir = Path(args.install_dir).resolve()
    preserve = json.loads(args.preserve or "[]")
    tmp = Path(tempfile.mkdtemp(prefix="douyin_update_"))
    zip_path = tmp / "update.zip"
    extract_dir = tmp / "extract"
    try:
        download_with_retry(args.download_url, zip_path)
        got = sha256_file(zip_path).lower()
        want = args.sha256.lower().strip()
        if not want:
            raise RuntimeError("version.json 缺少更新包 SHA256，已拒绝安装。")
        if got != want:
            raise RuntimeError(f"SHA256 mismatch. want={want}, got={got}")
        extract_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as z:
            validate_zip_members(z)
            z.extractall(extract_dir)
        children = [p for p in extract_dir.iterdir() if p.name != "__MACOSX"]
        if len(children) == 1 and children[0].is_dir() and not (extract_dir / "app").exists():
            src_root = children[0]
        else:
            src_root = extract_dir
        wait_process_exit(args.current_pid, seconds=20)
        copy_tree(src_root, install_dir, preserve)
        remove_obsolete_paths(install_dir)
        relaunch_without_console(install_dir, args.launch)
    except Exception as e:
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("更新失败", repr(e))
            root.destroy()
        except Exception:
            pass
        raise
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
