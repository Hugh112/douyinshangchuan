"""BitBrowser Local API adapter; never changes profiles, fingerprints or cookies."""

import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path


DEFAULT_BITBROWSER_API_URL = "http://127.0.0.1:54345"


class BitBrowserAPIError(RuntimeError):
    pass


def normalize_bitbrowser_api_token(value):
    value = str(value or "").strip()
    if len(value) > 4096 or not value.isascii() or "\r" in value or "\n" in value:
        raise BitBrowserAPIError("比特 API Token 格式无效，请重新复制完整 Token。")
    return value


class BitBrowserTokenStore:
    """Per-API credentials, encrypted for the current Windows user with DPAPI."""
    def __init__(self, path):
        self.path = Path(path)

    def _load(self):
        if not self.path.exists():
            return {}
        try:
            from platform_auth import _dpapi
            value = json.loads(_dpapi(self.path.read_bytes(), protect=False).decode("utf-8"))
            if not isinstance(value, dict) or value.get("format") != "bitbrowser-api-tokens-v1" or not isinstance(value.get("tokens"), dict):
                raise ValueError("invalid credential format")
            return value["tokens"]
        except Exception as exc:
            # Preserve the original file on corruption/decryption failure.
            raise BitBrowserAPIError("本机比特 API Token 无法解密读取，原文件已保留。") from exc

    def get_token(self, api_url):
        return normalize_bitbrowser_api_token(self._load().get(normalize_bitbrowser_api_url(api_url), ""))

    def save_token(self, api_url, token):
        from platform_auth import _dpapi
        api_url = normalize_bitbrowser_api_url(api_url)
        token = normalize_bitbrowser_api_token(token)
        tokens = self._load()
        if token:
            tokens[api_url] = token
        else:
            tokens.pop(api_url, None)
        if not tokens and not self.path.exists():
            return
        value = {"format": "bitbrowser-api-tokens-v1", "tokens": tokens}
        encrypted = _dpapi(json.dumps(value, ensure_ascii=False).encode("utf-8"), protect=True)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_name(self.path.name + ".tmp")
        temporary.write_bytes(encrypted)
        os.replace(temporary, self.path)


class _NoAPIRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


def normalize_bitbrowser_api_url(value):
    value = str(value or DEFAULT_BITBROWSER_API_URL).strip().rstrip("/")
    try:
        parsed = urllib.parse.urlsplit(value)
        valid_port = parsed.port is None or 1 <= parsed.port <= 65535
    except ValueError as exc:
        raise BitBrowserAPIError("比特浏览器 API 地址的端口无效。") from exc
    if (
        parsed.scheme not in {"http", "https"} or not parsed.hostname or not valid_port
        or parsed.username is not None or parsed.password is not None
        or parsed.query or parsed.fragment or parsed.path not in {"", "/"}
    ):
        raise BitBrowserAPIError("请填写比特浏览器系统设置中的本地接口地址，例如 http://127.0.0.1:54345。")
    host = parsed.hostname.casefold()
    if ":" in host:
        host = f"[{host}]"
    port = parsed.port
    suffix = f":{port}" if port and port != (443 if parsed.scheme == "https" else 80) else ""
    return f"{parsed.scheme}://{host}{suffix}"


def validate_bitbrowser_window_id(value):
    value = str(value or "").strip()
    if not value or any(char.isspace() for char in value) or value.isdecimal():
        raise BitBrowserAPIError("请填写比特浏览器窗口 ID（配置中的“复制 ID”），不能使用窗口序号。")
    return value


def bitbrowser_cdp_endpoint(data):
    if not isinstance(data, dict):
        raise BitBrowserAPIError("比特浏览器打开窗口后未返回调试连接数据。")
    for key in ("ws", "http"):
        value = str(data.get(key) or "").strip()
        if not value:
            continue
        if key == "http" and "://" not in value:
            value = "http://" + value
        try:
            parsed = urllib.parse.urlsplit(value)
            valid = parsed.scheme in ({"ws", "wss"} if key == "ws" else {"http", "https"})
            valid = valid and bool(parsed.hostname) and bool(parsed.port)
            valid = valid and parsed.username is None and parsed.password is None
        except ValueError:
            valid = False
        if valid:
            return value
    raise BitBrowserAPIError("比特浏览器未返回有效 ws/http 调试地址；请选择 Chrome 内核窗口并检查本地接口。")


class BitBrowserClient:
    def __init__(self, api_url=DEFAULT_BITBROWSER_API_URL, timeout=20, api_token=""):
        self.api_url = normalize_bitbrowser_api_url(api_url)
        self.timeout = timeout
        self.api_token = normalize_bitbrowser_api_token(api_token)
        # Local API must not be routed through system/browser proxy settings.
        self._opener = urllib.request.build_opener(urllib.request.ProxyHandler({}), _NoAPIRedirect())

    def request(self, path, payload=None, timeout=None):
        headers = {"Content-Type": "application/json", "Accept": "application/json"}
        if self.api_token:
            headers["x-api-key"] = self.api_token
        request = urllib.request.Request(
            self.api_url + path,
            data=json.dumps(payload or {}, ensure_ascii=False).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        try:
            with self._opener.open(request, timeout=timeout or self.timeout) as response:
                result = json.loads(response.read().decode("utf-8-sig"))
        except urllib.error.HTTPError as exc:
            if exc.code in {401, 403}:
                raise BitBrowserAPIError("比特浏览器 API 鉴权失败，请在账号设置中填写当前 Local API Token。") from exc
            raise BitBrowserAPIError(f"比特浏览器接口 {path} 返回 HTTP {exc.code}。") from exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            raise BitBrowserAPIError(
                f"无法连接比特浏览器接口 {path}；请启动并登录比特浏览器，核对系统设置中的本地接口地址。"
            ) from exc
        except (ValueError, UnicodeError) as exc:
            raise BitBrowserAPIError(f"比特浏览器接口 {path} 返回了无效 JSON。") from exc
        if not isinstance(result, dict) or result.get("success") is not True:
            message = str(result.get("msg") or "接口未返回 success=true") if isinstance(result, dict) else "返回格式无效"
            if self.api_token:
                message = message.replace(self.api_token, "[Token 已隐藏]")
            raise BitBrowserAPIError(f"比特浏览器接口 {path} 失败：{message[:300]}")
        return result.get("data")

    def list_windows(self):
        self.request("/health")
        windows = []
        seen = set()
        for page in range(1000):
            data = self.request("/browser/list", {"page": page, "pageSize": 100})
            if not isinstance(data, dict) or not isinstance(data.get("list"), list):
                raise BitBrowserAPIError("比特浏览器窗口列表格式无效，缺少 data.list。")
            items = data["list"]
            fresh_count = 0
            for item in items:
                if not isinstance(item, dict) or not item.get("id"):
                    raise BitBrowserAPIError("比特浏览器窗口列表含无效窗口 ID。")
                window_id = validate_bitbrowser_window_id(item["id"])
                if window_id in seen:
                    continue
                seen.add(window_id)
                fresh_count += 1
                # Do not retain platform passwords, cookies or proxy credentials.
                windows.append({
                    "id": window_id, "name": str(item.get("name") or ""),
                    "seq": item.get("seq", ""), "groupName": str(item.get("groupName") or ""),
                })
            if len(items) < 100:
                return windows
            if fresh_count == 0:
                raise BitBrowserAPIError("比特浏览器分页返回了重复数据，请重试读取窗口列表。")
            try:
                total = int(data.get("total"))
            except (ValueError, TypeError):
                total = None
            if total is not None and len(seen) >= total:
                return windows
        raise BitBrowserAPIError("比特浏览器窗口列表超过读取上限，请分批配置窗口 ID。")

    def window_is_open(self, window_id):
        window_id = validate_bitbrowser_window_id(window_id)
        data = self.request("/browser/pids/alive", {"ids": [window_id]})
        if not isinstance(data, dict):
            raise BitBrowserAPIError("比特浏览器进程查询格式无效，无法确认窗口是否由本软件启动。")
        try:
            return int(data.get(window_id) or 0) > 0
        except (TypeError, ValueError) as exc:
            raise BitBrowserAPIError("比特浏览器返回了无效的窗口进程 ID。") from exc

    def open_window(self, window_id, creator_url, reuse_existing=True, headless=False):
        window_id = validate_bitbrowser_window_id(window_id)
        already_open = self.window_is_open(window_id)
        if already_open and not reuse_existing:
            raise BitBrowserAPIError("该比特浏览器窗口已打开，请启用复用现有浏览器连接，或先手动关闭该窗口。")
        if already_open:
            # 查询现有端口即可连接，不再调用 open 消耗打开次数或新增默认标签。
            return {
                "browser_type": "bitbrowser", "api_url": self.api_url,
                "window_id": window_id, "cdp_endpoint": self.running_window_endpoint(window_id),
                "launched": False, "reused": True,
            }
        payload = {
            "id": window_id, "queue": True,
            "args": ["--headless=new", "--window-size=1440,1100"] if headless else [],
            "ignoreDefaultUrls": True, "newPageUrl": creator_url,
        }
        # Opening is never automatically retried: a timeout may still have started a window.
        data = self.request("/browser/open", payload, timeout=90)
        try:
            endpoint = bitbrowser_cdp_endpoint(data)
        except BitBrowserAPIError:
            try:
                endpoint = self.running_window_endpoint(window_id)
            except BitBrowserAPIError as exc:
                raise BitBrowserAPIError("比特窗口已启动但未返回可用调试地址；窗口已保留，不会自动关闭重开，请检查接口后重试连接。") from exc
        return {
            "browser_type": "bitbrowser", "api_url": self.api_url,
            "window_id": window_id, "cdp_endpoint": endpoint,
            "launched": not already_open, "reused": already_open,
        }

    def running_window_endpoint(self, window_id):
        data = self.request("/browser/ports")
        port = str(data.get(window_id) or "") if isinstance(data, dict) else ""
        if not port.isascii() or not port.isdigit() or len(port) > 5 or not 1 <= int(port) <= 65535:
            raise BitBrowserAPIError("比特浏览器已打开窗口的调试端口不可用；不会自动关闭重开，请检查窗口状态后重试连接。")
        host = urllib.parse.urlsplit(self.api_url).hostname
        host = f"[{host}]" if ":" in host else host
        return f"http://{host}:{int(port)}"

    def close_window(self, window_id):
        self.request("/browser/close", {"id": validate_bitbrowser_window_id(window_id)})
        # Official API requires 5 seconds before the same profile can be reopened.
        time.sleep(5)
