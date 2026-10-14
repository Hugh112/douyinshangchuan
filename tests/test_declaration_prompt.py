import sys
import unittest
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "app"))
import app_main as app


class DeclarationPromptTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.manager = sync_playwright().start()
        cls.browser = cls.manager.chromium.launch(channel="chrome", headless=True)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.manager.stop()

    def setUp(self):
        self.page = self.browser.new_page()
        self.page.set_content("""
        <div id="echo">自主声明：</div>
        <button id="publish" onclick="document.body.dataset.submitted='yes'">发布</button>
        <button id="outside" onclick="document.body.dataset.outside='yes'">确定</button>
        <section role="dialog" style="position:fixed;left:100px;top:100px;background:white;padding:30px;z-index:5">
          <h2>对作品内容添加声明</h2><p>请选择声明类型（单选）</p>
          <label><input type="radio" name="decl" value="ai">内容由AI生成</label>
          <label><input type="radio" name="decl" value="none" onchange="document.querySelector('#confirm').disabled=false">无需添加自主声明</label>
          <button id="confirm" disabled onclick="document.body.dataset.choice=document.querySelector('input:checked').value;document.getElementById('echo').textContent='自主声明：无需添加自主声明';this.closest('section').remove()">确定</button>
        </section>
        """)

    def tearDown(self):
        self.page.close()

    def test_selects_none_confirms_and_does_not_touch_other_buttons(self):
        self.assertTrue(app.handle_optional_declaration_prompt({}, self.page))
        self.assertEqual(self.page.locator("body").get_attribute("data-choice"), "none")
        self.assertIsNone(self.page.locator("body").get_attribute("data-submitted"))
        self.assertIsNone(self.page.locator("body").get_attribute("data-outside"))
        self.assertTrue(app.declaration_selected(self.page, "自主声明", "无需添加自主声明"))
        self.assertFalse(app.handle_optional_declaration_prompt({}, self.page))

    def test_other_dialog_with_similar_option_is_untouched(self):
        self.page.locator("h2").evaluate("el=>el.textContent='音乐设置'")
        self.page.locator("p").evaluate("el=>el.textContent='请选择音乐'")
        self.assertFalse(app.handle_optional_declaration_prompt({}, self.page))
        self.assertEqual(self.page.locator("input:checked").count(), 0)
        self.assertEqual(self.page.get_by_role("dialog").count(), 1)

    def test_inline_declaration_control_is_not_a_popup(self):
        self.page.get_by_role("dialog").evaluate("el=>{el.removeAttribute('role');el.style.position='static'}")
        self.assertFalse(app.handle_optional_declaration_prompt({}, self.page))
        self.assertEqual(self.page.locator("input:checked").count(), 0)

    def test_short_no_declaration_label(self):
        self.page.locator("label").last.evaluate("el=>el.lastChild.textContent='无需声明'")
        self.assertTrue(app.handle_optional_declaration_prompt({}, self.page))
        self.assertEqual(self.page.locator("body").get_attribute("data-choice"), "none")


if __name__ == "__main__":
    unittest.main()
