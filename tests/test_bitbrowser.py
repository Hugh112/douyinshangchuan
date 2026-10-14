import json
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "app"))

import app_main as app
from bitbrowser_api import BitBrowserAPIError, BitBrowserClient, BitBrowserTokenStore, bitbrowser_cdp_endpoint, normalize_bitbrowser_api_url


class LocalAPITests(unittest.TestCase):
    def setUp(self):
        self.calls = []
        self.api_keys = []
        self.reply = lambda path, body: {"success": True}
        test = self

        class Handler(BaseHTTPRequestHandler):
            def do_POST(self):
                body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
                test.calls.append((self.path, body, self.headers.get("Content-Type")))
                test.api_keys.append(self.headers.get("x-api-key"))
                result = test.reply(self.path, body)
                raw = json.dumps(result).encode("utf-8") if not isinstance(result, bytes) else result
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)

            def log_message(self, *_args):
                pass

        self.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.client = BitBrowserClient(f"http://127.0.0.1:{self.server.server_port}")

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()

    def test_json_post_pagination_and_sensitive_fields_discarded(self):
        def reply(path, body):
            if path == "/health":
                return {"success": True}
            page = body["page"]
            items = [{"id": f"window-{i}", "seq": i, "name": f"窗口{i}", "password": "PRIVATE", "cookie": "PRIVATE"}
                     for i in range(page * 100, min(102, (page + 1) * 100))]
            return {"success": True, "data": {"list": items, "total": 102}}

        self.reply = reply
        windows = self.client.list_windows()
        self.assertEqual(len(windows), 102)
        self.assertEqual([call[1] for call in self.calls[1:]], [{"page": 0, "pageSize": 100}, {"page": 1, "pageSize": 100}])
        self.assertEqual(windows[-1]["id"], "window-101")
        self.assertTrue(all(call[2] == "application/json" for call in self.calls))
        self.assertNotIn("PRIVATE", json.dumps(windows))

    def test_repeated_page_is_rejected_instead_of_truncated(self):
        items = [{"id": f"window-{i}"} for i in range(100)]
        self.reply = lambda path, body: {"success": True, "data": {"list": items}} if path != "/health" else {"success": True}
        with self.assertRaisesRegex(BitBrowserAPIError, "重复"):
            self.client.list_windows()
        self.assertEqual(len(self.calls), 3)

    def test_new_window_uses_returned_ws_and_closes_by_id(self):
        def reply(path, body):
            data = {} if path.endswith("alive") else {"ws": "ws://127.0.0.1:53325/devtools/browser/test", "http": "127.0.0.1:53325"}
            return {"success": True, "data": data}

        self.reply = reply
        info = self.client.open_window("window-A", "https://creator.douyin.com/")
        self.assertTrue(info["launched"])
        self.assertFalse(info["reused"])
        self.assertEqual(info["cdp_endpoint"], "ws://127.0.0.1:53325/devtools/browser/test")
        self.assertEqual(self.calls[1][1], {
            "id": "window-A", "queue": True, "args": [], "ignoreDefaultUrls": True,
            "newPageUrl": "https://creator.douyin.com/",
        })
        with patch("bitbrowser_api.time.sleep") as sleep:
            self.client.close_window("window-A")
            sleep.assert_called_once_with(5)
        self.assertEqual(self.calls[-1][0:2], ("/browser/close", {"id": "window-A"}))
        self.assertTrue(all(call[0] in {"/browser/open", "/browser/close", "/browser/pids/alive"} for call in self.calls))

    def test_existing_window_reused_and_never_marked_owned(self):
        self.reply = lambda path, body: {"success": True, "data": {"window-A": 123} if path.endswith("alive") else {"window-A": "53325"}}
        info = self.client.open_window("window-A", "https://creator.douyin.com/")
        self.assertFalse(info["launched"])
        self.assertTrue(info["reused"])
        self.assertEqual(info["cdp_endpoint"], f"http://127.0.0.1:53325")
        self.assertEqual([call[0] for call in self.calls], ["/browser/pids/alive", "/browser/ports"])
        with patch.object(app, "BitBrowserClient") as client, patch.object(app.subprocess, "run") as run:
            self.assertFalse(app.close_browser_launched_by_app(Mock(), info))
            client.assert_not_called()
            run.assert_not_called()

    def test_existing_window_without_reuse_does_not_open_or_close(self):
        self.reply = lambda path, body: {"success": True, "data": {"window-A": 123}}
        with self.assertRaisesRegex(BitBrowserAPIError, "已打开"):
            self.client.open_window("window-A", "https://creator.douyin.com/", reuse_existing=False)
        self.assertEqual(len(self.calls), 1)

    def test_business_error_and_invalid_json_do_not_retry_open(self):
        self.reply = lambda path, body: {"success": True, "data": {}} if path.endswith("alive") else {"success": False, "msg": "内核下载失败"}
        with self.assertRaisesRegex(BitBrowserAPIError, "内核下载失败"):
            self.client.open_window("window-A", "https://creator.douyin.com/")
        self.assertEqual(len(self.calls), 2)
        self.reply = lambda path, body: b"<html>invalid</html>"
        with self.assertRaisesRegex(BitBrowserAPIError, "JSON"):
            self.client.request("/health")

    def test_token_header_and_error_redaction(self):
        self.client = BitBrowserClient(self.client.api_url, api_token="fake-api-token-for-test")
        self.client.request("/health")
        self.assertEqual(self.api_keys[-1], "fake-api-token-for-test")
        self.reply = lambda path, body: {"success": False, "msg": "invalid token fake-api-token-for-test"}
        with self.assertRaises(BitBrowserAPIError) as error:
            self.client.request("/browser/list")
        self.assertNotIn("fake-api-token-for-test", str(error.exception))
        self.assertIn("已隐藏", str(error.exception))

    def test_without_token_does_not_send_auth_header(self):
        self.client.request("/health")
        self.assertIsNone(self.api_keys[-1])

    def test_token_is_not_forwarded_on_redirect(self):
        from urllib.response import addinfourl
        import io
        self.client = BitBrowserClient(self.client.api_url, api_token="fake-api-token-for-test")
        response = addinfourl(io.BytesIO(b""), {"Location": "http://example.org/"}, self.client.api_url, 302)
        response.msg = "Found"
        with patch.object(self.client._opener, "_open", return_value=response), self.assertRaises(BitBrowserAPIError):
            self.client.request("/health")

    def test_invalid_endpoint_never_closes_or_reopens_windows(self):
        for already_open in (False, True):
            self.calls.clear()
            self.reply = lambda path, body: {"success": True, "data": ({"window-A": 123} if already_open else {}) if path.endswith("alive") else {}}
            with self.assertRaisesRegex(BitBrowserAPIError, "调试"):
                self.client.open_window("window-A", "https://creator.douyin.com/")
            self.assertFalse(any(call[0] == "/browser/close" for call in self.calls))
            self.assertEqual(sum(call[0] == "/browser/open" for call in self.calls), int(not already_open))

    def test_new_window_missing_endpoint_recovers_from_ports_without_restart(self):
        self.reply = lambda path, body: {"success": True, "data": {"window-A": "53325"} if path.endswith("ports") else {}}
        info = self.client.open_window("window-A", "https://creator.douyin.com/")
        self.assertTrue(info["launched"])
        self.assertEqual(info["cdp_endpoint"], "http://127.0.0.1:53325")
        self.assertEqual([call[0] for call in self.calls], ["/browser/pids/alive", "/browser/open", "/browser/ports"])

    def test_repeated_existing_window_connections_use_no_open_calls(self):
        self.reply = lambda path, body: {"success": True, "data": {"window-A": 123} if path.endswith("alive") else {"window-A": "53325"}}
        for _ in range(3):
            self.assertTrue(self.client.open_window("window-A", "https://creator.douyin.com/")["reused"])
        self.assertEqual(sum(call[0] == "/browser/open" for call in self.calls), 0)
        self.reply = lambda path, body: {"success": True, "data": {"window-A": "65536"}}
        with self.assertRaisesRegex(BitBrowserAPIError, "端口不可用"):
            self.client.running_window_endpoint("window-A")


class TokenStorageTests(unittest.TestCase):
    def test_dpapi_encryption_url_isolation_and_clear(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "bitbrowser_api_tokens.bin"
            store = BitBrowserTokenStore(path)
            store.save_token("http://127.0.0.1:54345", "fake-token-for-54345")
            store.save_token("http://127.0.0.1:54346", "fake-token-for-54346")
            self.assertNotIn(b"fake-token", path.read_bytes())
            self.assertFalse(path.with_name(path.name + ".tmp").exists())
            self.assertEqual(store.get_token("http://127.0.0.1:54345/"), "fake-token-for-54345")
            self.assertEqual(store.get_token("http://127.0.0.1:54346"), "fake-token-for-54346")
            store.save_token("http://127.0.0.1:54345", "")
            self.assertEqual(store.get_token("http://127.0.0.1:54345"), "")
            self.assertEqual(store.get_token("http://127.0.0.1:54346"), "fake-token-for-54346")

    def test_corrupt_credential_file_is_preserved(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "bitbrowser_api_tokens.bin"
            path.write_bytes(b"not-a-dpapi-blob")
            with self.assertRaisesRegex(BitBrowserAPIError, "原文件已保留"):
                BitBrowserTokenStore(path).get_token("http://127.0.0.1:54345")
            self.assertEqual(path.read_bytes(), b"not-a-dpapi-blob")

    def test_api_token_not_retained_in_account_or_export_schema(self):
        account = app.normalize_browser_account({
            "browser_type": "bitbrowser", "bitbrowser_window_id": "test-window",
            "bitbrowser_api_token": "fake-token-not-for-config",
        }, 0, app.default_config())
        self.assertNotIn("fake-token", json.dumps(account))
        self.assertNotIn("bitbrowser_api_token", [key for _, key in app.BROWSER_QUEUE_TABLE_COLUMNS])

class BrowserIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.bit = app.normalize_browser_account({
            "id": "queue-bit", "browser_type": "bitbrowser", "name": "测试比特",
            "bitbrowser_api_url": "http://127.0.0.1:54345", "bitbrowser_window_id": "window-A",
            "browser_path": "", "cdp_port": 9222, "excel_path": "copy.xlsx", "sheet_name": "sheet1",
        }, 0, app.default_config())

    def test_legacy_identity_progress_and_mixed_queue_ports(self):
        legacy = app.default_config()
        original = {"id": "existing-id", "browser_path": r"D:\existing.lnk", "name": "已有账号", "cdp_port": 9222}
        chrome = app.normalize_browser_account(original, 0, legacy)
        self.assertEqual(chrome["browser_type"], "shortcut")
        self.assertEqual(chrome["id"], original["id"])
        before = app.config_for_browser_account(legacy, chrome)
        mixed = {**legacy, "browser_type": "bitbrowser", "bitbrowser_window_id": "window-A", "browser_accounts": [self.bit, chrome]}
        after = app.browser_accounts_from_config(mixed)[1]
        self.assertEqual(after["browser_type"], "shortcut")
        self.assertEqual(app.config_for_browser_account(mixed, after)["state_path"], before["state_path"])
        app.validate_browser_queue_ports([chrome, self.bit, {**self.bit, "id": "other", "bitbrowser_window_id": "window-B"}])
        with self.assertRaises(RuntimeError):
            app.validate_browser_queue_ports([chrome, {**chrome, "browser_path": r"D:\other.lnk"}])

    def test_bit_window_identity_and_independent_queue_progress(self):
        other = {**self.bit, "id": "repeat-row", "cdp_port": 9999}
        self.assertEqual(app.browser_runtime_identity(self.bit), app.browser_runtime_identity(other))
        self.assertNotEqual(app.browser_runtime_identity(self.bit), app.browser_runtime_identity({**other, "bitbrowser_window_id": "window-B"}))
        base = {"state_path": r"D:\progress.json"}
        first = app.config_for_browser_account(base, self.bit)
        second = app.config_for_browser_account(base, other)
        self.assertNotEqual(first["state_path"], second["state_path"])
        self.assertEqual(first["browser_type"], "bitbrowser")
        app.validate_browser_account_connection(self.bit, check_path=True)
        with self.assertRaisesRegex(BitBrowserAPIError, "序号"):
            app.validate_browser_account_connection({**self.bit, "bitbrowser_window_id": "35"})

    def test_dispatch_uses_bit_api_without_chrome_process_operations(self):
        cfg = app.config_for_browser_account(app.default_config(), self.bit)
        response = {"browser_type": "bitbrowser", "api_url": cfg["bitbrowser_api_url"], "window_id": "window-A", "launched": True, "reused": False, "cdp_endpoint": "ws://127.0.0.1:55555/devtools/browser/test"}
        with patch.object(app, "BitBrowserClient") as client, patch.object(app, "kill_chrome_residue") as kill, patch.object(app.subprocess, "Popen") as popen:
            client.return_value.open_window.return_value = response
            actual = app.open_browser_with_cdp(cfg)
            self.assertEqual(app.browser_connection_endpoint(cfg, actual), response["cdp_endpoint"])
            kill.assert_not_called()
            popen.assert_not_called()
            client.return_value.open_window.assert_called_once_with("window-A", cfg["creator_url"], reuse_existing=True, headless=False)
        with patch.object(app, "BitBrowserClient") as client, patch.object(app.subprocess, "run") as run:
            browser = Mock()
            self.assertTrue(app.close_browser_launched_by_app(browser, response))
            client.return_value.close_window.assert_called_once_with("window-A")
            run.assert_not_called()
            browser.close.assert_not_called()

    def test_mixed_queue_residue_cleanup_does_not_kill_bit_windows(self):
        with patch.object(app.subprocess, "run") as run:
            app.kill_chrome_residue({"close_chrome_before_start": True, "_protect_bitbrowser_windows": True})
            run.assert_not_called()

    def test_bit_tabs_keep_one_publish_page_and_reuse_it(self):
        workbench, creator, other = Mock(), Mock(), Mock()
        workbench.url, creator.url, other.url = "https://console.bitbrowser.net/?test=1", "https://creator.douyin.com/", "https://example.org/"
        workbench.is_closed.return_value = True
        other.is_closed.return_value = True
        context = Mock()
        context.pages = [workbench, creator, other]
        with patch.object(app, "ACTIVE_CFG", {"browser_type": "bitbrowser"}), patch.object(app, "register_page_health_monitor", side_effect=lambda page: page):
            self.assertIs(app.keep_single_browser_tab(context, other), creator)
            workbench.close.assert_called_once_with()
            creator.close.assert_not_called()
            other.close.assert_called_once_with()
            context.pages = [creator]
            self.assertIs(app.keep_single_browser_tab(context, creator), creator)
            context.new_page.assert_not_called()
            other.close.reset_mock()
            context.pages = [workbench, other]
            self.assertIs(app.keep_single_browser_tab(context, other), other)
            other.close.assert_not_called()
            context.new_page.assert_not_called()

    def test_bit_task_finish_closes_only_the_selected_window_by_api(self):
        self.assertTrue(self.bit["close_after_finish"])
        with patch.object(app, "bitbrowser_client") as client, patch.object(app.subprocess, "run") as run:
            client.return_value.window_is_open.return_value = True
            self.assertTrue(app.close_bitbrowser_task_window(self.bit))
            client.return_value.close_window.assert_called_once_with("window-A")
            run.assert_not_called()
        with patch.object(app, "bitbrowser_client") as client:
            client.return_value.window_is_open.return_value = False
            self.assertTrue(app.close_bitbrowser_task_window(self.bit))
            client.return_value.close_window.assert_not_called()

    def test_bit_queue_retains_window_until_its_last_task_finishes(self):
        other = {**self.bit, "id": "repeat-row"}
        self.assertTrue(app.bitbrowser_window_has_pending_accounts(self.bit, [self.bit, other], {self.bit["id"]}))
        self.assertFalse(app.bitbrowser_window_has_pending_accounts(self.bit, [self.bit, other], {self.bit["id"], other["id"]}))

    def test_bit_workbench_match_does_not_close_similar_or_local_user_pages(self):
        for url in ("http://127.0.0.1:54367/", "https://example.org/workbench/", "https://console.bitbrowser.net.example.org/", "https://console.bitbrowser.net/help", "https://creator.douyin.com/"):
            page = Mock(url=url)
            self.assertFalse(app.is_bitbrowser_workbench_page(page), url)

    def test_bit_workbench_close_failure_blocks_publish_tab(self):
        workbench = Mock(url="https://console.bitbrowser.net/")
        workbench.close.side_effect = RuntimeError("page close failed")
        context = Mock(pages=[workbench])
        with patch.object(app, "ACTIVE_CFG", {"browser_type": "bitbrowser"}), patch.object(app, "register_page_health_monitor") as monitor:
            with self.assertRaisesRegex(RuntimeError, "无法关闭.*工作台"):
                app.keep_single_browser_tab(context)
            context.new_page.assert_called_once_with()
            monitor.assert_not_called()

    def test_endpoint_and_url_validation(self):
        self.assertEqual(normalize_bitbrowser_api_url(" HTTP://LOCALHOST:54345/ "), "http://localhost:54345")
        self.assertEqual(bitbrowser_cdp_endpoint({"ws": "invalid", "http": "127.0.0.1:5000"}), "http://127.0.0.1:5000")
        for address in ("ftp://localhost:54345", "http://user:password@localhost:54345", "http://localhost:99999", "http://localhost:54345?key=secret"):
            with self.assertRaises(BitBrowserAPIError):
                normalize_bitbrowser_api_url(address)
        with self.assertRaises(ValueError):
            app.normalize_browser_type("unknown-browser")

    def test_publish_page_waits_for_async_module_loading(self):
        page = Mock()
        with patch.object(app, "creator_login_required", return_value=False), patch.object(app, "platform_image_publish_page_ready", side_effect=[False, False, True]), patch.object(app.time, "monotonic", side_effect=[0, 0, 1]):
            self.assertTrue(app.wait_for_image_publish_page({}, page))
        self.assertEqual(page.wait_for_timeout.call_count, 2)

    def test_publish_page_wait_is_bounded_and_detects_login(self):
        page = Mock()
        with patch.object(app, "creator_login_required", return_value=False), patch.object(app, "platform_image_publish_page_ready", return_value=False), patch.object(app.time, "monotonic", side_effect=[0, 0, 2]):
            self.assertFalse(app.wait_for_image_publish_page({"creator_page_ready_timeout_seconds": 1}, page))
        with patch.object(app, "creator_login_required", return_value=True):
            with self.assertRaises(app.AccountLoggedOutError):
                app.wait_for_image_publish_page({}, page)


class BitQueueLifecycleTests(unittest.TestCase):
    def run_queue(self, folder, accounts, rotate=False, fail=False):
        cfg = {**app.default_config(), "state_path": str(Path(folder) / "state.json"),
               "browser_accounts": accounts, "rotate_accounts_each_post": rotate, "rotate_batch_size": 1}
        path = Path(folder) / "config.json"
        path.write_text(json.dumps(cfg), encoding="utf-8")
        expected_config = path.read_bytes()
        events = []

        def child(command, **kwargs):
            child_cfg = json.loads(Path(command[-1]).read_text(encoding="utf-8"))
            account = child_cfg["browser_accounts"][0]
            self.assertFalse(account["close_after_finish"])
            resolved = app.config_for_browser_account(child_cfg, account)
            state = app.read_state(resolved)
            before = app.account_run_resume_progress(state, child_cfg["_queue_run_id"], account["posts_per_run"])
            events.append(("work", account["id"], before))
            increment = 1 if fail else min(account["posts_per_run"] - before, child_cfg.get("_queue_pass_limit", account["posts_per_run"]))
            state.update({"queue_run_id": child_cfg["_queue_run_id"], "run_progress": before + increment,
                          "last_account_outcome": "error" if fail else "done", "last_account_message": "fixture failure" if fail else ""})
            app.write_state(resolved, state)
            return Mock(stdout=[], wait=Mock(return_value=0))

        def close(account):
            events.append(("close", account["bitbrowser_window_id"]))
            return True

        with patch.object(app.subprocess, "Popen", side_effect=child), patch.object(app, "close_bitbrowser_task_window", side_effect=close):
            if fail:
                with self.assertRaisesRegex(RuntimeError, "进度已保留"):
                    app.worker(path)
            else:
                app.worker(path)
        self.assertEqual(path.read_bytes(), expected_config)
        return events

    def account(self, queue_id, window="window-A", quota=2):
        return app.normalize_browser_account({"id": queue_id, "name": queue_id, "browser_type": "bitbrowser",
                                              "bitbrowser_window_id": window, "posts_per_run": quota}, 0, {})

    def test_duplicate_rows_close_once_after_both_tasks(self):
        with tempfile.TemporaryDirectory() as folder:
            events = self.run_queue(folder, [self.account("row-a"), self.account("row-b", quota=1)])
            self.assertEqual(events, [("work", "row-a", 0), ("work", "row-b", 0), ("close", "window-A")])

    def test_rotation_keeps_windows_between_batches_and_closes_at_quota(self):
        with tempfile.TemporaryDirectory() as folder:
            events = self.run_queue(folder, [self.account("row-a"), self.account("row-b", "window-B")], rotate=True)
            self.assertEqual(events, [("work", "row-a", 0), ("work", "row-b", 0), ("work", "row-a", 1),
                                      ("close", "window-A"), ("work", "row-b", 1), ("close", "window-B")])

    def test_failed_task_closes_window_and_next_run_resumes_progress(self):
        with tempfile.TemporaryDirectory() as folder:
            accounts = [self.account("row-a")]
            first = self.run_queue(folder, accounts, fail=True)
            second = self.run_queue(folder, accounts)
            self.assertEqual(first, [("work", "row-a", 0), ("close", "window-A")])
            self.assertEqual(second, [("work", "row-a", 1), ("close", "window-A")])


if __name__ == "__main__":
    unittest.main()
