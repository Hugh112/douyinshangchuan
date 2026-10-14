import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "app"))
import app_main as app
from bitbrowser_api import BitBrowserTokenStore


class BrowserDialogTests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        self.token_path = Path(self.folder.name) / "bitbrowser_api_tokens.bin"
        self.root = app.tk.Tk()
        self.root.withdraw()
        self.token_patch = patch.object(app, "BITBROWSER_TOKEN_PATH", self.token_path)
        self.token_patch.start()
        self.copy_path = Path(self.folder.name) / "copy.xlsx"
        workbook = app.Workbook()
        workbook.active.append(["文案"])
        workbook.active.append(["测试内容"])
        workbook.save(self.copy_path)
        workbook.close()

    def tearDown(self):
        self.root.destroy()
        self.token_patch.stop()
        self.folder.cleanup()

    def dialog(self):
        account = app.normalize_browser_account({
            "id": "keep-queue-id", "name": "本地测试窗口", "browser_type": "bitbrowser",
            "bitbrowser_window_id": "test-window", "bitbrowser_api_url": "http://127.0.0.1:54345",
            "excel_path": str(self.copy_path), "sheet_name": "Sheet", "use_schedule": False,
        }, 0, {})
        with patch.object(app.tk.Toplevel, "wait_window"):
            return app.BrowserAccountDialog(self.root, account)

    def test_save_preserves_queue_id_and_keeps_token_out_of_account(self):
        dialog = self.dialog()
        dialog.vars["bitbrowser_api_token"].set("fake-dialog-token")
        with patch.object(app.messagebox, "showwarning") as warning:
            dialog.save_button.invoke()
            warning.assert_not_called()
        self.assertEqual(dialog.result["id"], "keep-queue-id")
        self.assertEqual(dialog.result["browser_type"], "bitbrowser")
        self.assertEqual(dialog.result["browser_path"], "")
        self.assertNotIn("bitbrowser_api_token", dialog.result)
        self.assertEqual(BitBrowserTokenStore(self.token_path).get_token("http://127.0.0.1:54345"), "fake-dialog-token")

    def test_api_address_change_does_not_send_token_to_another_server(self):
        store = BitBrowserTokenStore(self.token_path)
        store.save_token("http://127.0.0.1:54345", "fake-token-one")
        store.save_token("http://127.0.0.1:54346", "fake-token-two")
        dialog = self.dialog()
        self.assertEqual(dialog.vars["bitbrowser_api_token"].get(), "fake-token-one")
        dialog.vars["bitbrowser_api_url"].set("http://127.0.0.1:54346")
        self.assertEqual(dialog.vars["bitbrowser_api_token"].get(), "fake-token-two")
        dialog.vars["bitbrowser_api_url"].set("http://127.0.0.1:54347")
        self.assertEqual(dialog.vars["bitbrowser_api_token"].get(), "")

    def test_invalid_window_does_not_save_token_or_accept_dialog(self):
        dialog = self.dialog()
        dialog.vars["bitbrowser_window_id"].set("35")
        dialog.vars["bitbrowser_api_token"].set("fake-dialog-token")
        with patch.object(app.messagebox, "showwarning") as warning:
            dialog._save()
            warning.assert_called_once()
        self.assertIsNone(dialog.result)
        self.assertFalse(self.token_path.exists())

    def test_type_switch_and_token_field_masking(self):
        dialog = self.dialog()
        self.assertEqual(dialog.bitbrowser_form.winfo_manager(), "pack")
        self.assertTrue(dialog.close_var.get())
        self.assertFalse(dialog.close_checkbox.instate(["disabled"]))
        self.assertIn("比特窗口", dialog.close_checkbox.cget("text"))
        self.assertEqual(dialog.shortcut_form.winfo_manager(), "")
        token_entries = [child for child in dialog.bitbrowser_form.winfo_children()
                         if isinstance(child, app.ttk.Entry) and str(child.cget("textvariable")) == str(dialog.vars["bitbrowser_api_token"])]
        self.assertEqual(len(token_entries), 1)
        self.assertEqual(token_entries[0].cget("show"), "*")
        dialog.vars["browser_type"].set("快捷方式 / EXE")
        self.assertEqual(dialog.shortcut_form.winfo_manager(), "pack")
        self.assertEqual(dialog.bitbrowser_form.winfo_manager(), "")
        self.assertFalse(dialog.close_checkbox.instate(["disabled"]))

    def test_save_and_cancel_remain_visible_in_small_window_on_all_tabs(self):
        self.root.deiconify()
        dialog = self.dialog()
        dialog.window.geometry("640x420+20+20")
        self.root.update()
        for browser_type in ("快捷方式 / EXE", "比特浏览器"):
            dialog.vars["browser_type"].set(browser_type)
            for index in range(3):
                dialog.notebook.select(index)
                self.root.update()
                for button in (dialog.save_button, dialog.cancel_button):
                    self.assertTrue(button.winfo_ismapped())
                    self.assertGreaterEqual(button.winfo_rooty(), dialog.window.winfo_rooty())
                    self.assertLessEqual(button.winfo_rooty() + button.winfo_height(), dialog.window.winfo_rooty() + dialog.window.winfo_height())
                self.assertEqual(len(dialog.tab_scroll_canvases[index].yview()), 2)

    def test_stop_cleanup_tracks_only_invoked_bit_window_and_deduplicates_rows(self):
        gui = app.App.__new__(app.App)
        gui.proc_mode = "worker"
        bit = app.normalize_browser_account({"id": "bit-a", "browser_type": "bitbrowser", "bitbrowser_window_id": "window-A"}, 0, {})
        gui.browser_accounts = [bit, {**bit, "id": "bit-repeat"}, {**bit, "id": "not-invoked", "bitbrowser_window_id": "window-B"}]
        gui.write_ui = Mock()
        gui.remember_bitbrowser_task_account("bit-a")
        gui.remember_bitbrowser_task_account("bit-repeat")
        self.assertEqual(len(gui.active_bitbrowser_task_accounts), 1)
        # 暂停期间编辑了队列，仍按本次实际调用的原 ID 关闭。
        gui.browser_accounts[0]["bitbrowser_window_id"] = "changed-later"
        with patch.object(app, "close_bitbrowser_task_window", return_value=True) as close:
            gui.close_invoked_bitbrowser_task_windows()
            self.assertEqual(close.call_count, 1)
            self.assertEqual(close.call_args.args[0]["bitbrowser_window_id"], "window-A")
        self.assertEqual(gui.active_bitbrowser_task_accounts, {})

    def test_probe_does_not_register_publish_task_close_targets(self):
        gui = app.App.__new__(app.App)
        gui.proc_mode = "probe"
        gui.browser_accounts = []
        gui.remember_bitbrowser_task_account("some-window")
        self.assertFalse(getattr(gui, "active_bitbrowser_task_accounts", {}))


if __name__ == "__main__":
    unittest.main()
