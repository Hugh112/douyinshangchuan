# -*- coding: utf-8 -*-
"""Local-only BitBrowser connection form; does not publish or alter the account queue."""
import json
import os
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

from bitbrowser_api import (
    BitBrowserClient, BitBrowserTokenStore, DEFAULT_BITBROWSER_API_URL,
    normalize_bitbrowser_api_token, normalize_bitbrowser_api_url,
    validate_bitbrowser_window_id,
)
from app_main import BitBrowserWindowDialog


DATA_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "DouyinPublisher"
TOKEN_PATH = DATA_DIR / "bitbrowser_api_tokens.bin"
CONNECTION_PATH = DATA_DIR / "bitbrowser_connection_test.json"


def load_connection(path=CONNECTION_PATH):
    if not Path(path).is_file():
        return {}
    try:
        value = json.loads(Path(path).read_text(encoding="utf-8"))
        if isinstance(value, dict):
            return {key: str(value.get(key) or "") for key in ("api_url", "window_id", "window_name")}
    except (OSError, ValueError):
        pass
    return {}


def save_connection(api_url, token, window_id="", window_name="", path=CONNECTION_PATH, token_path=TOKEN_PATH):
    api_url = normalize_bitbrowser_api_url(api_url)
    token = normalize_bitbrowser_api_token(token)
    window_id = validate_bitbrowser_window_id(window_id) if str(window_id or "").strip() else ""
    value = {"api_url": api_url, "window_id": window_id, "window_name": str(window_name or "").strip()}
    BitBrowserTokenStore(token_path).save_token(api_url, token)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)
    return value


class BitBrowserSetupWindow:
    def __init__(self, root):
        self.root = root
        root.title("比特浏览器 API 本地配置与测试")
        root.geometry("780x450")
        root.minsize(720, 430)
        root.option_add("*Font", ("Microsoft YaHei UI", 10))
        self.token_store = BitBrowserTokenStore(TOKEN_PATH)
        saved = load_connection()
        self.api_url = tk.StringVar(value=saved.get("api_url") or DEFAULT_BITBROWSER_API_URL)
        self.api_token = tk.StringVar()
        self.window_id = tk.StringVar(value=saved.get("window_id", ""))
        self.window_name = tk.StringVar(value=saved.get("window_name", ""))
        self.status = tk.StringVar(value="填写本地接口地址；开启鉴权时填写 Token，然后读取窗口。")
        self._token_url = ""
        self._load_token()
        self.api_url.trace_add("write", self._load_token)
        self._build()

    def _build(self):
        outer = ttk.Frame(self.root, padding=24)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)
        ttk.Label(outer, text="比特浏览器连接配置", font=("Microsoft YaHei UI", 18, "bold")).grid(row=0, column=0, sticky="w")
        viewport = ttk.Frame(outer)
        viewport.grid(row=1, column=0, sticky="nsew", pady=(8, 12))
        viewport.columnconfigure(0, weight=1)
        viewport.rowconfigure(0, weight=1)
        canvas = tk.Canvas(viewport, highlightthickness=0, width=1, height=1)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(viewport, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)
        body = ttk.Frame(canvas)
        item = canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(item, width=event.width))
        self.root.bind("<MouseWheel>", lambda event: canvas.yview_scroll(-1 if event.delta > 0 else 1, "units") if str(event.widget).startswith(str(canvas)) else None, add="+")
        ttk.Label(body, text="先启动并登录比特浏览器；在测试窗口内登录抖音创作者中心。", wraplength=650).pack(anchor="w", pady=(0, 18))
        form = ttk.Frame(body)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        for row, (label, variable, masked) in enumerate((
            ("本地 API 地址", self.api_url, False),
            ("API Token（可空）", self.api_token, True),
            ("测试窗口 ID", self.window_id, False),
            ("窗口名称", self.window_name, False),
        )):
            ttk.Label(form, text=label, width=19).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=8)
            ttk.Entry(form, textvariable=variable, show="*" if masked else "").grid(row=row, column=1, sticky="ew", pady=8)
        actions = ttk.Frame(outer)
        actions.grid(row=2, column=0, sticky="ew")
        self.buttons = []
        for text, command in (("读取并选择窗口", self.select_window), ("测试连接", self.test_connection), ("保存设置", self.save)):
            button = ttk.Button(actions, text=text, command=command)
            button.pack(side="left", padx=(0, 12))
            self.buttons.append(button)
        ttk.Label(body, textvariable=self.status, wraplength=650).pack(anchor="w", pady=6)
        ttk.Label(body, text="Token 仅在本机加密保存；测试只读取窗口信息，不上传图片、不提交作品。", wraplength=650).pack(anchor="w", pady=(8, 0))

    def _load_token(self, *_args):
        try:
            api_url = normalize_bitbrowser_api_url(self.api_url.get())
        except Exception:
            self._token_url = ""
            self.api_token.set("")
            return
        if api_url == self._token_url:
            return
        self._token_url = api_url
        try:
            self.api_token.set(self.token_store.get_token(api_url))
        except Exception as exc:
            self.api_token.set("")
            self.status.set(str(exc))

    def _fetch(self, select):
        try:
            client = BitBrowserClient(self.api_url.get(), api_token=self.api_token.get())
            requested_id = self.window_id.get().strip()
            if requested_id and not select:
                validate_bitbrowser_window_id(requested_id)
        except Exception as exc:
            self.status.set(str(exc))
            return
        for button in self.buttons:
            button.configure(state="disabled")
        self.status.set("正在连接本地 API 并读取窗口列表…")
        results = queue.Queue()

        def fetch():
            try:
                results.put((client.list_windows(), None))
            except Exception as exc:
                results.put((None, str(exc)))

        def poll():
            try:
                windows, error = results.get_nowait()
            except queue.Empty:
                self.root.after(100, poll)
                return
            for button in self.buttons:
                button.configure(state="normal")
            if error:
                self.status.set(error)
                return
            self.status.set(f"连接成功：读取到 {len(windows)} 个窗口。")
            if select:
                if not windows:
                    self.status.set("API 连接成功，但当前账号没有窗口；请先在比特浏览器创建测试窗口。")
                    return
                chooser = BitBrowserWindowDialog(self.root, windows)
                if chooser.result:
                    self.window_id.set(chooser.result["id"])
                    self.window_name.set(chooser.result["name"])
                    self.status.set("已选择测试窗口，点击“保存设置”后告诉我即可。")
            elif requested_id:
                matching = next((item for item in windows if item["id"] == requested_id), None)
                if matching:
                    self.window_name.set(matching["name"])
                    self.status.set("API 鉴权与测试窗口查询成功。此检查尚未验证抖音页面。")
                else:
                    self.status.set("API 连接成功，但当前账号的窗口列表中没有该 ID，请重新读取并选择窗口。")

        threading.Thread(target=fetch, daemon=True).start()
        self.root.after(100, poll)

    def select_window(self):
        self._fetch(select=True)

    def test_connection(self):
        self._fetch(select=False)

    def save(self):
        try:
            save_connection(self.api_url.get(), self.api_token.get(), self.window_id.get(), self.window_name.get())
            self.status.set("设置已保存在本机。填写完成后告诉我，我再继续实机测试。")
        except Exception as exc:
            messagebox.showwarning("设置未保存", str(exc), parent=self.root)


if __name__ == "__main__":
    root = tk.Tk()
    BitBrowserSetupWindow(root)
    root.mainloop()
