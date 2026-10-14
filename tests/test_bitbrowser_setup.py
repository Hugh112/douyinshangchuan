import importlib.util
from importlib.machinery import SourceFileLoader
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

APP_DIR = Path(__file__).resolve().parents[1] / "app"
sys.path.insert(0, str(APP_DIR))
loader = SourceFileLoader("bitbrowser_setup", str(APP_DIR / "bitbrowser_setup.pyw"))
spec = importlib.util.spec_from_loader(loader.name, loader)
setup = importlib.util.module_from_spec(spec)
loader.exec_module(setup)


class ConnectionFormTests(unittest.TestCase):
    def test_save_separates_token_from_connection_metadata(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "connection.json"
            token_path = Path(folder) / "tokens.bin"
            setup.save_connection("http://127.0.0.1:54345/", "fake-private-token", "test-window", "测试窗口", path, token_path)
            self.assertEqual(setup.load_connection(path), {
                "api_url": "http://127.0.0.1:54345", "window_id": "test-window", "window_name": "测试窗口",
            })
            self.assertNotIn("fake-private-token", path.read_text(encoding="utf-8"))
            self.assertNotIn(b"fake-private-token", token_path.read_bytes())
            self.assertEqual(setup.BitBrowserTokenStore(token_path).get_token("http://127.0.0.1:54345"), "fake-private-token")
            self.assertEqual(set(json.loads(path.read_text(encoding="utf-8"))), {"api_url", "window_id", "window_name"})

    def test_invalid_window_does_not_write_connection_or_credentials(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "connection.json"
            token_path = Path(folder) / "tokens.bin"
            with self.assertRaisesRegex(RuntimeError, "序号"):
                setup.save_connection("http://127.0.0.1:54345", "fake-private-token", "35", "", path, token_path)
            self.assertFalse(path.exists())
            self.assertFalse(token_path.exists())

    def test_gui_has_masked_token_and_can_be_closed_without_saving(self):
        root = setup.tk.Tk()
        root.withdraw()
        try:
            with patch.object(setup, "load_connection", return_value={}), patch.object(setup, "BitBrowserTokenStore") as store, patch.object(setup, "save_connection") as save:
                store.return_value.get_token.return_value = ""
                form = setup.BitBrowserSetupWindow(root)
                self.assertEqual(len(form.buttons), 3)
                self.assertEqual(root.title(), "比特浏览器 API 本地配置与测试")
                entries = []
                def collect(widget):
                    for child in widget.winfo_children():
                        if isinstance(child, setup.ttk.Entry):
                            entries.append(child)
                        collect(child)
                collect(root)
                self.assertEqual(len(entries), 4)
                self.assertEqual(sum(child.cget("show") == "*" for child in entries), 1)
                save.assert_not_called()
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
