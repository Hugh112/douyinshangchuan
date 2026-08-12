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

UPDATER_VERSION = "3.0.6"


def format_bytes(value):
    try:
        size = max(0, int(value or 0))
    except Exception:
        size = 0
    units = ("B", "KB", "MB", "GB")
    amount = float(size)
    for unit in units:
        if amount < 1024 or unit == units[-1]:
            return f"{amount:.0f} {unit}" if unit == "B" else f"{amount:.1f} {unit}"
        amount /= 1024
    return f"{size} B"


def stage_percent(start, end, completed, total):
    try:
        total = int(total or 0)
        completed = max(0, int(completed or 0))
    except Exception:
        return float(start)
    if total <= 0:
        return float(start)
    ratio = min(1.0, completed / total)
    return float(start) + (float(end) - float(start)) * ratio


class UpdateProgressWindow:
    """同步更新流程使用的小型进度窗口；Tk 不可用时静默退化为无界面模式。"""

    def __init__(self, target_version=""):
        self.root = None
        self.progress = None
        self.status_var = None
        self.detail_var = None
        self.percent_var = None
        self._last_pump = 0.0
        try:
            import tkinter as tk
            from tkinter import ttk

            root = tk.Tk()
            root.title("抖音智能发布中心 · 正在更新")
            root.geometry("540x230")
            root.resizable(False, False)
            root.protocol("WM_DELETE_WINDOW", lambda: None)
            root.configure(bg="#F3F6FA")
            try:
                resource_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
                icon_path = resource_root / "assets" / "app_logo.ico"
                if icon_path.exists():
                    root.iconbitmap(default=str(icon_path))
            except Exception:
                pass

            frame = tk.Frame(root, bg="#FFFFFF", bd=0, padx=26, pady=22)
            frame.pack(fill="both", expand=True, padx=14, pady=14)
            title = "正在更新抖音智能发布中心"
            if str(target_version or "").strip():
                title += f"  v{str(target_version).strip()}"
            tk.Label(
                frame, text=title, bg="#FFFFFF", fg="#172033",
                font=("Microsoft YaHei UI", 14, "bold"), anchor="w",
            ).pack(fill="x")

            self.status_var = tk.StringVar(value="正在准备更新…")
            self.detail_var = tk.StringVar(value="请不要关闭电脑或结束更新程序。")
            self.percent_var = tk.StringVar(value="0%")
            status_row = tk.Frame(frame, bg="#FFFFFF")
            status_row.pack(fill="x", pady=(20, 7))
            tk.Label(
                status_row, textvariable=self.status_var, bg="#FFFFFF", fg="#26344D",
                font=("Microsoft YaHei UI", 11, "bold"), anchor="w",
            ).pack(side="left")
            tk.Label(
                status_row, textvariable=self.percent_var, bg="#FFFFFF", fg="#2563EB",
                font=("Microsoft YaHei UI", 11, "bold"), anchor="e",
            ).pack(side="right")
            self.progress = ttk.Progressbar(frame, orient="horizontal", mode="determinate", maximum=100)
            self.progress.pack(fill="x", ipady=5)
            tk.Label(
                frame, textvariable=self.detail_var, bg="#FFFFFF", fg="#667085",
                font=("Microsoft YaHei UI", 9), anchor="w", justify="left", wraplength=475,
            ).pack(fill="x", pady=(11, 0))
            self.root = root
            root.lift()
            self.set_progress(1, "正在准备更新…", "正在创建安全的临时更新目录。", force=True)
        except Exception:
            self.root = None

    def _pump(self, force=False):
        if self.root is None:
            return
        now = time.monotonic()
        if not force and now - self._last_pump < 0.05:
            return
        self._last_pump = now
        try:
            self.root.update_idletasks()
            self.root.update()
        except Exception:
            self.root = None

    def set_progress(self, percent, status, detail="", force=False):
        if self.root is None:
            return
        value = max(0.0, min(100.0, float(percent or 0)))
        self.progress["value"] = value
        self.status_var.set(str(status or "正在更新…"))
        self.detail_var.set(str(detail or ""))
        self.percent_var.set(f"{int(round(value))}%")
        self._pump(force=force)

    def wait_tick(self, percent, status, detail):
        self.set_progress(percent, status, detail, force=True)

    def show_error(self, message):
        self.set_progress(
            self.progress["value"] if self.progress is not None else 0,
            "更新失败",
            str(message),
            force=True,
        )

    def close(self):
        if self.root is None:
            return
        try:
            self.root.destroy()
        except Exception:
            pass
        self.root = None


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


def download_with_retry(url, dst, retry=3, timeout=60, expected_size=0, progress_callback=None):
    parsed = urllib.parse.urlparse(str(url or ""))
    if parsed.scheme.lower() != "https":
        raise RuntimeError("更新包下载地址必须使用 HTTPS。")
    headers = {"User-Agent": f"DouyinPublisherUpdater/{UPDATER_VERSION}", "Cache-Control": "no-cache", "Pragma": "no-cache", "Connection": "close"}
    last_err = None
    sep = "&" if "?" in url else "?"
    url2 = f"{url}{sep}_t={int(time.time())}"
    for i in range(max(1, retry)):
        try:
            req = urllib.request.Request(url2, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as r, open(dst, "wb") as f:
                try:
                    total = int(r.headers.get("Content-Length") or expected_size or 0)
                except Exception:
                    total = int(expected_size or 0)
                downloaded = 0
                while True:
                    chunk = r.read(1024 * 1024)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        progress_callback(downloaded, total, i + 1, max(1, retry), None)
                if total > 0 and downloaded != total:
                    raise RuntimeError(f"更新包下载不完整：应为 {total} 字节，实际 {downloaded} 字节。")
            return
        except Exception as e:
            last_err = e
            if progress_callback:
                progress_callback(0, int(expected_size or 0), i + 1, max(1, retry), e)
            time.sleep(2 + i * 2)
    raise last_err


def sha256_file(path, progress_callback=None):
    h = hashlib.sha256()
    total = Path(path).stat().st_size
    completed = 0
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
            completed += len(chunk)
            if progress_callback:
                progress_callback(completed, total)
    return h.hexdigest()


def wait_process_exit(pid, seconds=8, progress_callback=None):
    try:
        pid = int(pid or 0)
    except Exception:
        pid = 0
    if not pid:
        return
    if os.name != "nt":
        deadline = time.monotonic() + min(2, max(0, float(seconds)))
        while time.monotonic() < deadline:
            if progress_callback:
                progress_callback(max(0, deadline - time.monotonic()))
            time.sleep(0.2)
        return
    try:
        import ctypes
        SYNCHRONIZE = 0x00100000
        handle = ctypes.windll.kernel32.OpenProcess(SYNCHRONIZE, False, pid)
        if handle:
            try:
                deadline = time.monotonic() + max(0, float(seconds))
                while time.monotonic() < deadline:
                    result = ctypes.windll.kernel32.WaitForSingleObject(handle, 200)
                    if result == 0:
                        break
                    if progress_callback:
                        progress_callback(max(0, deadline - time.monotonic()))
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        else:
            deadline = time.monotonic() + min(2, max(0, float(seconds)))
            while time.monotonic() < deadline:
                if progress_callback:
                    progress_callback(max(0, deadline - time.monotonic()))
                time.sleep(0.2)
    except Exception:
        deadline = time.monotonic() + min(2, max(0, float(seconds)))
        while time.monotonic() < deadline:
            if progress_callback:
                progress_callback(max(0, deadline - time.monotonic()))
            time.sleep(0.2)


def copy_tree(src, dst, preserve_paths, progress_callback=None):
    src = Path(src)
    dst = Path(dst)
    preserve_paths = set(str(p).replace(chr(92), "/") for p in preserve_paths or [])
    backups = {}
    for rel in preserve_paths:
        p = dst / rel
        if p.exists():
            backups[rel] = p.read_bytes()
    pending_files = []
    for root, dirs, files in os.walk(src):
        rootp = Path(root)
        rel_root = rootp.relative_to(src)
        for name in files:
            src_file = rootp / name
            rel = (rel_root / name).as_posix()
            if rel in preserve_paths:
                continue
            pending_files.append((src_file, rel))
    total = len(pending_files)
    for index, (src_file, rel) in enumerate(pending_files, 1):
        dst_file = dst / rel
        dst_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_file, dst_file)
        if progress_callback:
            progress_callback(index, total, rel)
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


def extract_zip_with_progress(archive, extract_dir, progress_callback=None):
    members = archive.infolist()
    total = len(members)
    for index, info in enumerate(members, 1):
        archive.extract(info, extract_dir)
        if progress_callback:
            progress_callback(index, total, info.filename)


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


def run_self_test():
    assert UPDATER_VERSION == "3.0.6"
    assert format_bytes(0) == "0 B"
    assert format_bytes(1024) == "1.0 KB"
    assert stage_percent(5, 60, 50, 100) == 32.5
    assert stage_percent(5, 60, 0, 0) == 5.0
    print("UPDATER_SELF_TEST_OK")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--install-dir", default="")
    ap.add_argument("--current-pid", default="0")
    ap.add_argument("--download-url", default="")
    ap.add_argument("--sha256", default="")
    ap.add_argument("--version", default="")
    ap.add_argument("--package-size", default="0")
    ap.add_argument("--launch", default="")
    ap.add_argument("--preserve", default="[]")
    args = ap.parse_args()
    if args.self_test:
        run_self_test()
        return
    if not args.install_dir or not args.download_url or not args.sha256:
        ap.error("--install-dir、--download-url 和 --sha256 为必填参数")

    install_dir = Path(args.install_dir).resolve()
    preserve = json.loads(args.preserve or "[]")
    tmp = Path(tempfile.mkdtemp(prefix="douyin_update_"))
    zip_path = tmp / "update.zip"
    extract_dir = tmp / "extract"
    try:
        expected_size = max(0, int(args.package_size or 0))
    except Exception:
        expected_size = 0
    progress = UpdateProgressWindow(args.version)
    try:
        progress.set_progress(3, "正在连接更新服务器…", "准备从统一云服务器下载更新包。", force=True)

        def on_download(completed, total, attempt, retries, error):
            if error is not None:
                progress.set_progress(
                    5,
                    f"下载中断，准备第 {min(attempt + 1, retries)}/{retries} 次尝试…",
                    str(error),
                    force=True,
                )
                return
            display_total = total or expected_size
            percent = stage_percent(5, 60, completed, display_total)
            detail = f"已下载 {format_bytes(completed)}"
            if display_total:
                detail += f" / {format_bytes(display_total)}"
            if retries > 1:
                detail += f"（第 {attempt}/{retries} 次尝试）"
            progress.set_progress(percent, "正在下载更新包…", detail)

        download_with_retry(
            args.download_url,
            zip_path,
            expected_size=expected_size,
            progress_callback=on_download,
        )
        if expected_size and zip_path.stat().st_size != expected_size:
            raise RuntimeError(
                f"更新包大小不一致：应为 {expected_size} 字节，实际 {zip_path.stat().st_size} 字节。"
            )

        progress.set_progress(61, "正在校验更新包…", "正在计算 SHA256，请稍候。", force=True)
        got = sha256_file(
            zip_path,
            progress_callback=lambda completed, total: progress.set_progress(
                stage_percent(61, 70, completed, total),
                "正在校验更新包…",
                f"已校验 {format_bytes(completed)} / {format_bytes(total)}",
            ),
        ).lower()
        want = args.sha256.lower().strip()
        if not want:
            raise RuntimeError("version.json 缺少更新包 SHA256，已拒绝安装。")
        if got != want:
            raise RuntimeError(f"SHA256 mismatch. want={want}, got={got}")

        extract_dir.mkdir(parents=True, exist_ok=True)
        progress.set_progress(71, "正在检查更新包…", "正在检查 ZIP 路径和文件安全性。", force=True)
        with zipfile.ZipFile(zip_path, "r") as z:
            validate_zip_members(z)
            extract_zip_with_progress(
                z,
                extract_dir,
                progress_callback=lambda completed, total, name: progress.set_progress(
                    stage_percent(72, 83, completed, total),
                    "正在解压更新包…",
                    f"{completed}/{total}  {str(name)[-90:]}",
                ),
            )
        children = [p for p in extract_dir.iterdir() if p.name != "__MACOSX"]
        if len(children) == 1 and children[0].is_dir() and not (extract_dir / "app").exists():
            src_root = children[0]
        else:
            src_root = extract_dir

        progress.set_progress(84, "正在等待主程序退出…", "主程序退出后将立即替换文件。", force=True)
        wait_process_exit(
            args.current_pid,
            seconds=20,
            progress_callback=lambda remaining: progress.wait_tick(
                84,
                "正在等待主程序退出…",
                f"仍在等待主程序释放文件，最多约 {int(remaining) + 1} 秒。",
            ),
        )
        progress.set_progress(85, "正在安装更新…", "开始替换程序文件，用户配置不会被覆盖。", force=True)
        copy_tree(
            src_root,
            install_dir,
            preserve,
            progress_callback=lambda completed, total, rel: progress.set_progress(
                stage_percent(85, 97, completed, total),
                "正在安装更新…",
                f"{completed}/{total}  {str(rel)[-90:]}",
            ),
        )
        progress.set_progress(98, "正在清理旧文件…", "只清理更新器明确列出的历史程序文件。", force=True)
        remove_obsolete_paths(install_dir)
        progress.set_progress(99, "正在重新启动…", "更新已经安装，正在无黑框启动新版本。", force=True)
        relaunch_without_console(install_dir, args.launch)
        progress.set_progress(100, "更新完成", "新版本已启动，本窗口即将自动关闭。", force=True)
        deadline = time.monotonic() + 1.2
        while time.monotonic() < deadline:
            progress._pump(force=True)
            time.sleep(0.1)
    except Exception as e:
        progress.show_error(repr(e))
        try:
            from tkinter import messagebox

            messagebox.showerror(
                "更新失败",
                f"更新没有完成，原有用户数据不会被删除。\n\n{repr(e)}",
                parent=progress.root,
            )
        except Exception:
            pass
        raise
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        progress.close()


if __name__ == "__main__":
    main()
