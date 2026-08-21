import sys
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPOSITORY_ROOT / "app"))

from platform_auth import (  # noqa: E402
    PlatformAuthClient,
    PlatformAuthError,
    STATE_AUTHORIZED,
    STATE_SIGNED_OUT,
    VALIDATION_LOGIN_ONLY,
)


def time_text(value):
    return value.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def login_envelope(now=None):
    now = now or datetime.now(timezone.utc)
    return {
        "data": {
            "user": {"id": "user-id", "account": "background-test"},
            "product": {
                "id": "product-id",
                "code": "publisher.douyin",
                "name": "抖音智能发布中心",
            },
            "license": {
                "id": "license-id",
                "licenseNumber": "TEST",
                "status": "Active",
                "startsAtUtc": time_text(now - timedelta(days=1)),
                "expiresAtUtc": time_text(now + timedelta(days=30)),
                "isPerpetual": False,
                "remainingDays": 30,
                "deviceLimit": 2,
                "concurrencyLimit": 1,
                "offlineGraceSeconds": 300,
            },
            "features": {"publish.enabled": True},
            "authentication": {
                "stage": "session-established",
                "databaseValidated": True,
                "sessionEstablished": True,
                "validationMode": "login_once_per_login",
            },
            "session": {
                "id": "session-id",
                "deviceId": "device-id",
                "licenseDeviceId": "binding-id",
                "accessToken": "access-token",
                "accessTokenExpiresAtUtc": time_text(now + timedelta(minutes=20)),
                "refreshToken": "refresh-token",
                "refreshTokenExpiresAtUtc": time_text(now + timedelta(days=30)),
            },
        },
        "serverTimeUtc": time_text(now),
    }


class MemoryStore:
    def __init__(self, value=None):
        self.value = value

    def load(self, product_code):
        if self.value and self.value.get("productCode") == product_code:
            return self.value
        return None

    def save(self, value):
        self.value = value

    def delete(self):
        self.value = None


def create_client(store=None):
    client = PlatformAuthClient(
        "https://platform.example.test/",
        "publisher.douyin",
        "3.0.6",
        REPOSITORY_ROOT / "tests" / "unused-token.bin",
    )
    client.store = store or MemoryStore()
    return client


class LoginOnlyAuthorizationTests(unittest.TestCase):
    def test_default_heartbeat_does_not_contact_server(self):
        client = create_client()
        requests = []

        def request(method, path, body=None, access_token="", timeout=None):
            requests.append((method, path))
            return login_envelope()

        client._request = request
        client.login("background-test", "memory-only")

        for _ in range(10):
            client.heartbeat()

        self.assertEqual(VALIDATION_LOGIN_ONLY, client.validation_mode)
        self.assertEqual([("POST", "api/v1/auth/login")], requests)
        self.assertEqual(STATE_AUTHORIZED, client.state)

    def test_explicit_and_implicit_refresh_require_next_login(self):
        client = create_client()
        requests = []

        def request(method, path, body=None, access_token="", timeout=None):
            requests.append((method, path))
            return login_envelope()

        client._request = request
        client.login("background-test", "memory-only")

        with self.assertRaises(PlatformAuthError) as explicit:
            client.refresh()
        self.assertEqual("SESSION_RELOGIN_REQUIRED", explicit.exception.code)

        client.session["accessTokenExpiresAtUtc"] = time_text(
            datetime.now(timezone.utc) - timedelta(seconds=1)
        )
        with self.assertRaises(PlatformAuthError) as implicit:
            client.get_release_policy()
        self.assertEqual("SESSION_RELOGIN_REQUIRED", implicit.exception.code)
        self.assertEqual([("POST", "api/v1/auth/login")], requests)
        self.assertEqual(STATE_AUTHORIZED, client.state)

    def test_saved_session_refreshes_once_as_next_login(self):
        now = datetime.now(timezone.utc)
        initial = login_envelope(now)["data"]
        snapshot = dict(initial)
        session = dict(snapshot.pop("session"))
        session.pop("accessToken")
        refresh_token = session.pop("refreshToken")
        snapshot.update(
            {
                "sessionId": session["id"],
                "deviceId": session["deviceId"],
                "licenseDeviceId": session["licenseDeviceId"],
                "accessTokenExpiresAtUtc": session["accessTokenExpiresAtUtc"],
                "refreshTokenExpiresAtUtc": session["refreshTokenExpiresAtUtc"],
                "lastServerTimeUtc": time_text(now),
                "lastSuccessfulContactUtc": time_text(now),
            }
        )
        store = MemoryStore(
            {
                "schemaVersion": 1,
                "baseUri": "https://platform.example.test/",
                "productCode": "publisher.douyin",
                "deviceFingerprintHash": "test-fingerprint",
                "refreshToken": refresh_token,
                "snapshot": snapshot,
            }
        )
        client = create_client(store)
        client.device = {
            "fingerprintHash": "test-fingerprint",
            "displayName": "test",
            "osName": "Windows",
            "osVersion": "11",
        }
        requests = []

        def request(method, path, body=None, access_token="", timeout=None):
            requests.append((method, path))
            refreshed = dict(login_envelope()["data"]["session"])
            return {"data": refreshed, "serverTimeUtc": time_text(datetime.now(timezone.utc))}

        client._request = request

        restored = client.auto_login()

        self.assertIsNotNone(restored)
        self.assertEqual([("POST", "api/v1/auth/refresh")], requests)
        self.assertEqual(STATE_AUTHORIZED, client.state)

    def test_retryable_restore_failure_does_not_use_offline_snapshot(self):
        now = datetime.now(timezone.utc)
        data = login_envelope(now)["data"]
        session = data["session"]
        snapshot = {
            "user": data["user"],
            "product": data["product"],
            "license": data["license"],
            "features": data["features"],
            "sessionId": session["id"],
            "deviceId": session["deviceId"],
            "licenseDeviceId": session["licenseDeviceId"],
            "accessTokenExpiresAtUtc": session["accessTokenExpiresAtUtc"],
            "refreshTokenExpiresAtUtc": session["refreshTokenExpiresAtUtc"],
            "lastServerTimeUtc": time_text(now),
            "lastSuccessfulContactUtc": time_text(now),
        }
        store = MemoryStore(
            {
                "schemaVersion": 1,
                "baseUri": "https://platform.example.test/",
                "productCode": "publisher.douyin",
                "deviceFingerprintHash": "test-fingerprint",
                "refreshToken": session["refreshToken"],
                "snapshot": snapshot,
            }
        )
        client = create_client(store)
        client.device = {
            "fingerprintHash": "test-fingerprint",
            "displayName": "test",
            "osName": "Windows",
            "osVersion": "11",
        }

        def unavailable(*_args, **_kwargs):
            raise PlatformAuthError(
                "AUTHORIZATION_UNREACHABLE",
                "temporary",
                retryable=True,
            )

        client._request = unavailable

        self.assertIsNone(client.auto_login())
        self.assertEqual(STATE_SIGNED_OUT, client.state)
        self.assertIsNone(client.session)
        self.assertIsNotNone(store.value)


if __name__ == "__main__":
    unittest.main()
