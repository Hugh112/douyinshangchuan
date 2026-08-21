# -*- coding: utf-8 -*-
"""统一软件授权平台客户端（Windows 桌面端）。

仅上传产品码、客户端版本和不可逆设备指纹。令牌使用当前 Windows 用户的
DPAPI 加密后保存；不会上传抖音账号、Cookie、正文、图片或本地路径。
"""

from __future__ import annotations

import ctypes
import hashlib
import json
import os
import platform
import socket
import threading
import urllib.error
import urllib.parse
import urllib.request
import uuid
from ctypes import wintypes
from datetime import datetime, timedelta, timezone
from pathlib import Path


STATE_SIGNED_OUT = "signed_out"
STATE_AUTHENTICATING = "authenticating"
STATE_AUTHORIZED = "authorized"
STATE_GRACE = "grace_period"
STATE_FORCE_LOGOUT = "force_logout"

VALIDATION_LOGIN_ONLY = "login_only"
VALIDATION_CONTINUOUS = "continuous"


class PlatformAuthError(RuntimeError):
    def __init__(self, code, message, retryable=False, status=0, trace_id=""):
        super().__init__(str(message or code or "授权平台请求失败"))
        self.code = str(code or "PLATFORM_REQUEST_FAILED")
        self.message = str(message or self.code)
        self.retryable = bool(retryable)
        self.status = int(status or 0)
        self.trace_id = str(trace_id or "")


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _blob_from_bytes(value):
    data = bytes(value or b"")
    buffer = ctypes.create_string_buffer(data, max(1, len(data)))
    blob = _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    return blob, buffer


def _dpapi(value, protect):
    """使用 CryptProtectData/CryptUnprotectData；返回的新缓冲由 LocalFree 释放。"""
    if os.name != "nt":
        raise PlatformError("DPAPI 只支持 Windows。")
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    source, source_buffer = _blob_from_bytes(value)
    output = _DataBlob()
    flags = 0x01  # CRYPTPROTECT_UI_FORBIDDEN
    if protect:
        ok = crypt32.CryptProtectData(
            ctypes.byref(source), None, None, None, None, flags, ctypes.byref(output)
        )
    else:
        ok = crypt32.CryptUnprotectData(
            ctypes.byref(source), None, None, None, None, flags, ctypes.byref(output)
        )
    _ = source_buffer
    if not ok:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(output.pbData, output.cbData)
    finally:
        kernel32.LocalFree(output.pbData)


class PlatformError(RuntimeError):
    pass


def _utc_now():
    return datetime.now(timezone.utc)


def _parse_time(value):
    if isinstance(value, datetime):
        return value.astimezone(timezone.utc)
    text = str(value or "").strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        result = datetime.fromisoformat(text)
    except ValueError:
        return None
    if result.tzinfo is None:
        result = result.replace(tzinfo=timezone.utc)
    return result.astimezone(timezone.utc)


def _time_text(value):
    parsed = _parse_time(value)
    return parsed.isoformat().replace("+00:00", "Z") if parsed else ""


def _read_machine_guid():
    if os.name == "nt":
        try:
            import winreg

            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Microsoft\Cryptography",
                0,
                winreg.KEY_READ | getattr(winreg, "KEY_WOW64_64KEY", 0),
            ) as key:
                value, _kind = winreg.QueryValueEx(key, "MachineGuid")
                if str(value or "").strip():
                    return str(value).strip()
        except OSError:
            pass
    return "|".join((socket.gethostname(), platform.machine(), str(os.cpu_count() or 1)))


def windows_device_identity():
    stable_source = _read_machine_guid().encode("utf-8")
    return {
        "fingerprintHash": hashlib.sha256(stable_source).hexdigest().upper(),
        "displayName": socket.gethostname()[:128],
        "osName": "Windows" if os.name == "nt" else platform.system(),
        "osVersion": platform.platform()[:128],
    }


class DpapiTokenStore:
    def __init__(self, path):
        self.path = Path(path).resolve()

    def load(self, product_code):
        if not self.path.is_file():
            return None
        encrypted = self.path.read_bytes()
        try:
            raw = _dpapi(encrypted, protect=False)
            value = json.loads(raw.decode("utf-8"))
        except Exception as exc:
            self.delete()
            raise PlatformError("本机授权会话无法读取，请重新登录。") from exc
        if not isinstance(value, dict) or value.get("productCode") != product_code:
            return None
        return value

    def save(self, value):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        raw = json.dumps(value, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        encrypted = _dpapi(raw, protect=True)
        temporary = self.path.with_name(self.path.name + ".tmp")
        temporary.write_bytes(encrypted)
        os.replace(temporary, self.path)

    def delete(self):
        try:
            self.path.unlink()
        except FileNotFoundError:
            pass


class PlatformAuthClient:
    """线程安全的登录、刷新、心跳、退出和发布策略客户端。"""

    def __init__(
        self,
        base_url,
        product_code,
        client_version,
        token_path,
        release_channel="stable",
        timeout=15,
        validation_mode=VALIDATION_LOGIN_ONLY,
    ):
        parsed = urllib.parse.urlparse(str(base_url or "").strip())
        is_loopback = parsed.hostname in {"127.0.0.1", "localhost", "::1"}
        if parsed.scheme != "https" and not (parsed.scheme == "http" and is_loopback):
            raise ValueError("生产授权 API 必须使用 HTTPS。")
        if not parsed.netloc:
            raise ValueError("授权 API 地址无效。")
        self.base_url = str(base_url).rstrip("/") + "/"
        self.product_code = str(product_code).strip()
        self.client_version = str(client_version).strip()
        self.release_channel = str(release_channel).strip() or "stable"
        self.validation_mode = str(validation_mode or VALIDATION_LOGIN_ONLY).strip().lower()
        if self.validation_mode not in {VALIDATION_LOGIN_ONLY, VALIDATION_CONTINUOUS}:
            raise ValueError("授权验证模式无效。")
        self.timeout = max(5, min(120, int(timeout or 15)))
        self.store = DpapiTokenStore(token_path)
        self.device = windows_device_identity()
        self.state = STATE_SIGNED_OUT
        self.session = None
        self.access_token = ""
        self.refresh_token = ""
        self.last_error = None
        self._lock = threading.RLock()
        self._opener = urllib.request.build_opener(_NoRedirect())

    @property
    def account(self):
        return str(((self.session or {}).get("user") or {}).get("account") or "")

    @property
    def can_start_new_task(self):
        features = (self.session or {}).get("features") or {}
        return self.state == STATE_AUTHORIZED and features.get("publish.enabled") is True

    @property
    def access_expires_at(self):
        return _parse_time((self.session or {}).get("accessTokenExpiresAtUtc"))

    def _request(self, method, path, body=None, access_token="", timeout=None):
        url = urllib.parse.urljoin(self.base_url, str(path).lstrip("/"))
        headers = {
            "Accept": "application/json",
            "User-Agent": f"DouyinPublisher/{self.client_version}",
            "Cache-Control": "no-cache",
        }
        data = None
        if body is not None:
            data = json.dumps(body, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
            headers["Content-Type"] = "application/json"
        if access_token:
            headers["Authorization"] = "Bearer " + str(access_token)
        request = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with self._opener.open(request, timeout=timeout or self.timeout) as response:
                raw = response.read()
                if response.status == 204:
                    return None
        except urllib.error.HTTPError as exc:
            raw = exc.read()
            try:
                payload = json.loads(raw.decode("utf-8-sig"))
                error = payload.get("error") or {}
            except Exception:
                payload, error = {}, {}
            raise PlatformAuthError(
                error.get("code") or f"HTTP_{exc.code}",
                error.get("message") or f"授权平台返回 HTTP {exc.code}",
                error.get("retryable", exc.code >= 500),
                exc.code,
                payload.get("traceId", ""),
            ) from exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            raise PlatformAuthError(
                "AUTHORIZATION_UNREACHABLE",
                "暂时无法连接授权服务器，请检查网络后重试。",
                retryable=True,
            ) from exc
        try:
            payload = json.loads(raw.decode("utf-8-sig"))
        except Exception as exc:
            raise PlatformAuthError(
                "PLATFORM_CONTRACT_INVALID", "授权平台返回的数据格式无效。"
            ) from exc
        if not isinstance(payload, dict) or "data" not in payload:
            raise PlatformAuthError(
                "PLATFORM_CONTRACT_INVALID", "授权平台响应缺少 data。"
            )
        return payload

    def _stored_value(self):
        session = dict(self.session or {})
        session.pop("accessToken", None)
        return {
            "schemaVersion": 1,
            "baseUri": self.base_url,
            "productCode": self.product_code,
            "deviceFingerprintHash": self.device["fingerprintHash"],
            "refreshToken": self.refresh_token,
            "snapshot": session,
        }

    def _save(self):
        self.store.save(self._stored_value())

    def _set_login_session(self, login_data, server_time):
        login_data = login_data if isinstance(login_data, dict) else {}
        session = dict(login_data.get("session") or {})
        self.access_token = str(session.pop("accessToken", "") or "")
        self.refresh_token = str(session.pop("refreshToken", "") or "")
        if not self.access_token or not self.refresh_token:
            raise PlatformAuthError(
                "PLATFORM_CONTRACT_INVALID", "登录响应缺少授权令牌。"
            )
        self.session = {
            "user": login_data.get("user") or {},
            "product": login_data.get("product") or {},
            "license": login_data.get("license") or {},
            "features": login_data.get("features") or {},
            "sessionId": session.get("id", ""),
            "deviceId": session.get("deviceId", ""),
            "licenseDeviceId": session.get("licenseDeviceId", ""),
            "accessTokenExpiresAtUtc": session.get("accessTokenExpiresAtUtc", ""),
            "refreshTokenExpiresAtUtc": session.get("refreshTokenExpiresAtUtc", ""),
            "lastServerTimeUtc": _time_text(server_time),
            "lastSuccessfulContactUtc": _time_text(_utc_now()),
        }
        product = self.session["product"]
        if product.get("code") != self.product_code:
            raise PlatformAuthError(
                "PRODUCT_MISMATCH", "授权响应中的产品码与本软件不一致。"
            )
        self.state = STATE_AUTHORIZED
        self.last_error = None
        self._save()

    def login(self, account, password):
        account, password = str(account or "").strip(), str(password or "")
        if not account or not password:
            raise PlatformAuthError("CREDENTIALS_REQUIRED", "请输入软件授权账号和密码。")
        with self._lock:
            self.state = STATE_AUTHENTICATING
            try:
                response = self._request(
                    "POST",
                    "api/v1/auth/login",
                    {
                        "productCode": self.product_code,
                        "account": account,
                        "password": password,
                        "clientVersion": self.client_version,
                        "releaseChannel": self.release_channel,
                        "device": self.device,
                    },
                )
                self._set_login_session(response["data"], response.get("serverTimeUtc"))
                return dict(self.session)
            except Exception as exc:
                self.state = STATE_SIGNED_OUT
                self.session = None
                self.access_token = self.refresh_token = ""
                self.last_error = exc
                raise

    def auto_login(self):
        with self._lock:
            stored = self.store.load(self.product_code)
            if not stored:
                self.state = STATE_SIGNED_OUT
                return None
            if stored.get("baseUri") != self.base_url:
                self.store.delete()
                self.state = STATE_SIGNED_OUT
                return None
            if stored.get("deviceFingerprintHash") != self.device["fingerprintHash"]:
                self.store.delete()
                self.state = STATE_SIGNED_OUT
                return None
            self.session = dict(stored.get("snapshot") or {})
            self.refresh_token = str(stored.get("refreshToken") or "")
            refresh_expiry = _parse_time(self.session.get("refreshTokenExpiresAtUtc"))
            if not self.refresh_token or not refresh_expiry or refresh_expiry <= _utc_now():
                self.clear_local_session(STATE_SIGNED_OUT)
                return None
            self.state = STATE_AUTHENTICATING
            try:
                return self._refresh_for_login()
            except PlatformAuthError as exc:
                if exc.retryable and self.validation_mode == VALIDATION_CONTINUOUS:
                    self._enter_grace(exc)
                    return dict(self.session) if self.session else None
                if exc.retryable:
                    # 登录单次验证模式下，启动恢复本身就是一次新登录。
                    # 本次在线验证失败时不能使用旧授权快照放行；保留受保护的
                    # Refresh Token 文件，便于下次启动再次在线验证。
                    self.state = STATE_SIGNED_OUT
                    self.session = None
                    self.access_token = self.refresh_token = ""
                    self.last_error = exc
                    return None
                self.clear_local_session(STATE_SIGNED_OUT)
                return None

    def _update_refreshed_session(self, session_data, server_time):
        session_data = dict(session_data or {})
        access_token = str(session_data.pop("accessToken", "") or "")
        refresh_token = str(session_data.pop("refreshToken", "") or "")
        if not access_token or not refresh_token:
            raise PlatformAuthError(
                "PLATFORM_CONTRACT_INVALID", "刷新响应缺少轮换后的授权令牌。"
            )
        self.access_token = access_token
        self.refresh_token = refresh_token
        self.session.update(
            {
                "sessionId": session_data.get("id", self.session.get("sessionId", "")),
                "deviceId": session_data.get("deviceId", self.session.get("deviceId", "")),
                "licenseDeviceId": session_data.get(
                    "licenseDeviceId", self.session.get("licenseDeviceId", "")
                ),
                "accessTokenExpiresAtUtc": session_data.get("accessTokenExpiresAtUtc", ""),
                "refreshTokenExpiresAtUtc": session_data.get("refreshTokenExpiresAtUtc", ""),
                "lastServerTimeUtc": _time_text(server_time),
                "lastSuccessfulContactUtc": _time_text(_utc_now()),
            }
        )
        self.state = STATE_AUTHORIZED
        self.last_error = None
        self._save()

    def _refresh_for_login(self):
        with self._lock:
            if not self.session or not self.refresh_token:
                raise PlatformAuthError("STORED_SESSION_NOT_FOUND", "没有可刷新的本机授权会话。")
            response = self._request(
                "POST",
                "api/v1/auth/refresh",
                {
                    "productCode": self.product_code,
                    "sessionId": self.session.get("sessionId", ""),
                    "deviceFingerprintHash": self.device["fingerprintHash"],
                    "refreshToken": self.refresh_token,
                },
            )
            self._update_refreshed_session(response["data"], response.get("serverTimeUtc"))
            return dict(self.session)

    def refresh(self):
        if self.validation_mode == VALIDATION_LOGIN_ONLY:
            raise PlatformAuthError(
                "SESSION_RELOGIN_REQUIRED",
                "当前登录已固定授权快照；退出后重新登录可再次验证授权。",
            )
        return self._refresh_for_login()

    def _ensure_access_token(self):
        expiry = self.access_expires_at
        if not self.access_token or not expiry or expiry <= _utc_now() + timedelta(seconds=45):
            if self.validation_mode == VALIDATION_LOGIN_ONLY:
                raise PlatformAuthError(
                    "SESSION_RELOGIN_REQUIRED",
                    "当前登录的访问令牌已到期；退出后重新登录可继续访问平台附加服务。",
                )
            self._refresh_for_login()

    def heartbeat(self):
        with self._lock:
            if not self.session:
                raise PlatformAuthError("SESSION_NOT_AVAILABLE", "当前没有有效授权会话。")
            if self.validation_mode == VALIDATION_LOGIN_ONLY:
                return dict(self.session)
            try:
                self._ensure_access_token()
                response = self._request(
                    "POST",
                    "api/v1/auth/heartbeat",
                    {"productCode": self.product_code},
                    self.access_token,
                )
                self.session["lastServerTimeUtc"] = _time_text(response.get("serverTimeUtc"))
                self.session["lastSuccessfulContactUtc"] = _time_text(_utc_now())
                self.state = STATE_AUTHORIZED
                self.last_error = None
                self._save()
                return dict(self.session)
            except PlatformAuthError as exc:
                self.last_error = exc
                if exc.retryable:
                    self._enter_grace(exc)
                    return dict(self.session) if self.session else None
                self.clear_local_session(STATE_FORCE_LOGOUT)
                raise

    def _enter_grace(self, error):
        if not self.session:
            self.state = STATE_FORCE_LOGOUT
            return
        now = _utc_now()
        last_contact = _parse_time(self.session.get("lastSuccessfulContactUtc"))
        license_data = self.session.get("license") or {}
        grace_seconds = max(0, int(license_data.get("offlineGraceSeconds") or 0))
        license_expiry = _parse_time(license_data.get("expiresAtUtc"))
        if not last_contact or now < last_contact - timedelta(minutes=5):
            self.clear_local_session(STATE_FORCE_LOGOUT)
            return
        if (
            grace_seconds <= 0
            or now > last_contact + timedelta(seconds=grace_seconds)
            or (not license_data.get("isPerpetual") and license_expiry and now >= license_expiry)
        ):
            self.clear_local_session(STATE_FORCE_LOGOUT)
            return
        self.state = STATE_GRACE
        self.last_error = error

    def get_release_policy(self, timeout=6):
        with self._lock:
            if self.state != STATE_AUTHORIZED:
                raise PlatformAuthError("AUTHORIZATION_REQUIRED", "需要先登录授权账号。")
            self._ensure_access_token()
            query = urllib.parse.urlencode(
                {
                    "productCode": self.product_code,
                    "channel": self.release_channel,
                    "currentVersion": self.client_version,
                }
            )
            response = self._request(
                "GET",
                "api/v1/client/release-policy?" + query,
                access_token=self.access_token,
                timeout=timeout,
            )
            return response.get("data") or {}

    def _report_telemetry(self, endpoint, body, idempotency_key=""):
        with self._lock:
            if self.state != STATE_AUTHORIZED:
                raise PlatformAuthError(
                    "AUTHORIZATION_REQUIRED", "只有在线授权状态可以登记客户端遥测。"
                )
            self._ensure_access_token()
            payload = {
                "productCode": self.product_code,
                "clientVersion": self.client_version,
                "idempotencyKey": str(idempotency_key or uuid.uuid4().hex)[:96],
                "occurredAtUtc": _time_text(_utc_now()),
            }
            payload.update(body)
            response = self._request(
                "POST",
                endpoint,
                payload,
                access_token=self.access_token,
            )
            return response.get("data") or {}

    def report_usage(self, metric_code, quantity=1, idempotency_key=""):
        return self._report_telemetry(
            "api/v1/client/telemetry/usage",
            {
                "metricCode": str(metric_code),
                "quantity": float(quantity),
            },
            idempotency_key,
        )

    def report_task(
        self,
        status,
        attempted_count,
        succeeded_count,
        failed_count,
        task_type="publish.batch",
        idempotency_key="",
    ):
        return self._report_telemetry(
            "api/v1/client/telemetry/task",
            {
                "taskType": str(task_type),
                "status": str(status),
                "attemptedCount": max(0, int(attempted_count or 0)),
                "succeededCount": max(0, int(succeeded_count or 0)),
                "failedCount": max(0, int(failed_count or 0)),
            },
            idempotency_key,
        )

    def report_error(
        self,
        error_code,
        stage,
        retryable=False,
        idempotency_key="",
    ):
        return self._report_telemetry(
            "api/v1/client/telemetry/error",
            {
                "errorCode": str(error_code),
                "stage": str(stage),
                "retryable": bool(retryable),
            },
            idempotency_key,
        )

    def logout(self):
        with self._lock:
            if self.access_token:
                try:
                    self._request(
                        "POST",
                        "api/v1/auth/logout",
                        {"productCode": self.product_code},
                        self.access_token,
                    )
                except Exception:
                    pass
            self.clear_local_session(STATE_SIGNED_OUT)

    def clear_local_session(self, state=STATE_SIGNED_OUT):
        self.store.delete()
        self.session = None
        self.access_token = self.refresh_token = ""
        self.state = state

    def next_heartbeat_delay(self, first_contact=False, configured_seconds=60):
        if self.validation_mode == VALIDATION_LOGIN_ONLY:
            return None
        if first_contact:
            return 10
        if self.state == STATE_GRACE:
            return 15
        online = max(5, min(60, int(configured_seconds or 60)))
        expiry = self.access_expires_at
        if not expiry:
            return online
        refresh_delay = max(5, int((expiry - _utc_now()).total_seconds() - 45))
        return min(online, refresh_delay)

    def status_text(self):
        if self.state == STATE_AUTHENTICATING:
            return "授权验证中"
        if self.state == STATE_AUTHORIZED:
            account = self.account or "已登录"
            license_data = (self.session or {}).get("license") or {}
            remaining = license_data.get("remainingDays")
            suffix = "永久授权" if license_data.get("isPerpetual") else (
                f"剩余 {remaining} 天" if remaining is not None else "授权有效"
            )
            return f"已授权：{account} · {suffix}"
        if self.state == STATE_GRACE:
            return "授权服务器暂不可达 · 宽限重连中"
        if self.state == STATE_FORCE_LOGOUT:
            return "授权已失效 · 请重新登录"
        return "未登录软件授权"
