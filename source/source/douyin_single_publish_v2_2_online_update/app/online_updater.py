# -*- coding: utf-8 -*-
"""
抖音单图文发布 - 在线更新器
从 GitHub 下载新版压缩包，校验 SHA256，保留本机配置后覆盖安装。
"""

import argparse
import ctypes
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path


def msgbox(title, text, style=0x40):
    try:
        ctypes.windll.user32.MessageBoxW(None, str(text), str(title), style | 0x00001000)
    except Exception:
        print(f"{title}: {text}")


def wait_pid(pid, timeout=30):
    if not pid or pid <= 0 or os.name != "nt":
        time.sleep(1.5)
        return
    # 轮询 Windows 进程是否退出，不依赖第三方库。
    for _ in range(timeout * 2):
        try:
            r = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}"],
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="ignore",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            out = r.stdout or ""
            if str(pid) not in out:
                return
        except Exception:
            pass
        time.sleep(0.5)


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def download_file(url, target):
    req = urllib.request.Request(url, headers={"User-Agent": "DouyinPublisher-Updater"})
    with urllib.request.urlopen(req, timeout=60) as resp, open(target, "wb") as f:
        shutil.copyfileobj(resp, f)


def find_package_root(extract_dir):
    items = [p for p in Path(extract_dir).iterdir() if p.name != "__MACOSX"]
    if len(items) == 1 and items[0].is_dir():
        return items[0]
    return Path(extract_dir)


def is_preserved(rel_path, preserve_paths):
    rel = rel_path.replace("\\", "/").strip("/")
    for item in preserve_paths:
        item = str(item).replace("\\", "/").strip("/")
        if not item:
            continue
        if rel == item or rel.startswith(item + "/"):
            return True
    return False


def copy_update(src_root, install_dir, preserve_paths):
    install_dir = Path(install_dir)
    src_root = Path(src_root)
    install_dir.mkdir(parents=True, exist_ok=True)

    backup_dir = install_dir.parent / (install_dir.name + "_update_backup_" + time.strftime("%Y%m%d_%H%M%S"))
    backup_dir.mkdir(parents=True, exist_ok=True)

    # 先备份需要保留的文件/目录，避免新版包覆盖本机配置。
    for rel in preserve_paths:
        rel_norm = str(rel).replace("\\", "/").strip("/")
        old = install_dir / rel_norm
        if old.exists():
            dst = backup_dir / rel_norm
            dst.parent.mkdir(parents=True, exist_ok=True)
            if old.is_dir():
                shutil.copytree(old, dst, dirs_exist_ok=True)
            else:
                shutil.copy2(old, dst)

    for item in src_root.rglob("*"):
        rel = item.relative_to(src_root)
        rel_text = str(rel).replace("\\", "/")
        if rel_text.startswith(".git/") or rel_text == ".git":
            continue
        if is_preserved(rel_text, preserve_paths):
            continue
        target = install_dir / rel
        if item.is_dir():
            target.mkdir(parents=True, exist_ok=True)
        else:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(item, target)

    # 恢复保留文件/目录。
    for rel in preserve_paths:
        rel_norm = str(rel).replace("\\", "/").strip("/")
        bak = backup_dir / rel_norm
        if bak.exists():
            dst = install_dir / rel_norm
            dst.parent.mkdir(parents=True, exist_ok=True)
            if bak.is_dir():
                shutil.copytree(bak, dst, dirs_exist_ok=True)
            else:
                shutil.copy2(bak, dst)

    return backup_dir


def launch_app(launch_path):
    launch_path = Path(launch_path)
    if not launch_path.exists():
        return
    exe = Path(sys.executable)
    # 有 pythonw 就优先静默启动。
    if exe.name.lower() == "python.exe":
        pyw = exe.with_name("pythonw.exe")
        if pyw.exists():
            exe = pyw
    subprocess.Popen([str(exe), str(launch_path)], cwd=str(launch_path.parent))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--install-dir", required=True)
    ap.add_argument("--current-pid", type=int, default=0)
    ap.add_argument("--download-url", required=True)
    ap.add_argument("--sha256", default="")
    ap.add_argument("--launch", default="")
    ap.add_argument("--preserve", default="[]")
    args = ap.parse_args()

    try:
        preserve_paths = json.loads(args.preserve or "[]")
        if not isinstance(preserve_paths, list):
            preserve_paths = ["app/douyin_gui_config.json"]
    except Exception:
        preserve_paths = ["app/douyin_gui_config.json"]

    try:
        wait_pid(args.current_pid)
        with tempfile.TemporaryDirectory(prefix="douyin_update_") as td:
            td = Path(td)
            zip_path = td / "update.zip"
            download_file(args.download_url, zip_path)
            expected = (args.sha256 or "").strip().lower()
            if expected:
                actual = sha256_file(zip_path)
                if actual != expected:
                    raise RuntimeError(f"更新包校验失败。\n期望：{expected}\n实际：{actual}")
            extract_dir = td / "extract"
            extract_dir.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(zip_path) as z:
                z.extractall(extract_dir)
            src_root = find_package_root(extract_dir)
            backup = copy_update(src_root, Path(args.install_dir), preserve_paths)
        msgbox("更新完成", "更新完成，已保留本机配置。\n备份目录：" + str(backup))
        if args.launch:
            launch_app(args.launch)
    except Exception as e:
        msgbox("更新失败", repr(e), 0x10)


if __name__ == "__main__":
    main()
