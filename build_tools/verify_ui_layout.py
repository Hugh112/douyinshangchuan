"""Check all application buttons for clipping using isolated Tk fixtures."""
import argparse
import importlib.util
import json
import sys
import tempfile
from importlib.machinery import SourceFileLoader
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / 'app'))
import app_main as app
loader = SourceFileLoader('layout_setup', str(PROJECT_ROOT / 'app' / 'bitbrowser_setup.pyw'))
spec = importlib.util.spec_from_loader(loader.name, loader)
setup = importlib.util.module_from_spec(spec)
loader.exec_module(setup)
BUTTON_TYPES = (app.ttk.Button, app.ttk.Checkbutton, app.tk.Button, app.tk.Checkbutton)


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def audit_scale(scaling):
    root = app.tk.Tk()
    root.withdraw()
    root.tk.call('tk', 'scaling', scaling)
    checked, issues = [], []

    def measure(window, button, label):
        name = str(button.cget('text'))
        checked.append({'surface': label, 'button': name})
        if not button.winfo_ismapped():
            issues.append([label, name, 'not mapped'])
            return
        canvas, parent = None, button.master
        while parent is not None:
            if isinstance(parent, app.tk.Canvas):
                canvas = parent
                break
            parent = getattr(parent, 'master', None)
        if canvas is not None:
            box = canvas.bbox('all')
            if box:
                y = canvas.canvasy(0) + button.winfo_rooty() - canvas.winfo_rooty()
                canvas.yview_moveto(max(0, (y + button.winfo_height() / 2 - canvas.winfo_height() / 2) / max(1, box[3])))
                window.update()
        target = canvas or window
        br = (button.winfo_rootx(), button.winfo_rooty(), button.winfo_width(), button.winfo_height())
        tr = (target.winfo_rootx(), target.winfo_rooty(), target.winfo_width(), target.winfo_height())
        if br[0] < tr[0] - 2 or br[1] < tr[1] - 2 or br[0] + br[2] > tr[0] + tr[2] + 2 or br[1] + br[3] > tr[1] + tr[3] + 2:
            issues.append([label, name, 'clipped', br, tr])
        if br[2] < button.winfo_reqwidth() - 2 or br[3] < button.winfo_reqheight() - 2:
            issues.append([label, name, 'compressed', br, [button.winfo_reqwidth(), button.winfo_reqheight()]])

    def surface(window, scope, label, excluded=()):
        for button in descendants(scope):
            if isinstance(button, BUTTON_TYPES) and not any(str(button).startswith(str(hidden)) for hidden in excluded):
                measure(window, button, label)

    try:
        with tempfile.TemporaryDirectory(prefix='douyin_layout_') as folder:
            cfg = app.default_config()
            cfg.update({'state_path': str(Path(folder) / 'state.json'), 'debug_dir': str(Path(folder) / 'screenshots')})
            with patch.object(app.App, 'start_authorization_restore', return_value=None), patch.object(app, 'load_config', return_value=cfg), patch.object(app, 'BITBROWSER_TOKEN_PATH', Path(folder) / 'bit.bin'), patch.object(app, 'AUTH_LOGIN_PREFERENCES_PATH', Path(folder) / 'prefs.bin'), patch.object(app, 'AUTH_TOKEN_PATH', Path(folder) / 'auth.bin'), patch.object(sys, 'argv', ['layout', '--self-test']):
                gui = app.App(root)
                root.minsize(1, 1)
                root.geometry('1080x660+20+20')
                root.deiconify()
                root.update()
                for index in range(4):
                    gui.notebook.select(index)
                    root.update()
                    surface(root, gui.notebook.nametowidget(gui.notebook.tabs()[index]), f'main-tab-{index}')
                for button in descendants(root):
                    if isinstance(button, BUTTON_TYPES) and not str(button).startswith(str(gui.notebook)):
                        measure(root, button, 'main-fixed')
                with patch.object(app.tk.Toplevel, 'wait_window'):
                    dialog = app.BrowserAccountDialog(root, app.normalize_browser_account({'browser_type': 'bitbrowser', 'bitbrowser_window_id': 'fixture-window'}, 0, {}), '编辑账号布局检查')
                dialog.window.geometry('640x420+20+20')
                root.update()
                for browser_type in ('比特浏览器', '快捷方式 / EXE'):
                    dialog.vars['browser_type'].set(browser_type)
                    for index in range(3):
                        dialog.notebook.select(index)
                        root.update()
                        hidden = dialog.shortcut_form if browser_type == '比特浏览器' else dialog.bitbrowser_form
                        surface(dialog.window, dialog.notebook.nametowidget(dialog.notebook.tabs()[index]), f'editor-{browser_type}-{index}', (hidden,))
                        for button in (dialog.save_button, dialog.cancel_button):
                            measure(dialog.window, button, 'editor-fixed')
                dialog.window.destroy()
                with patch.object(app.tk.Toplevel, 'wait_window'):
                    selector = app.BitBrowserWindowDialog(root, [{'seq': 1, 'name': 'fixture', 'groupName': 'fixture', 'id': 'fixture-window'}])
                selector.window.geometry('700x420+20+20')
                root.update()
                surface(selector.window, selector.window, 'window-selector')
                selector.window.destroy()
                gui.show_authorization_login()
                root.update()
                login = gui.authorization_window
                surface(login, login, 'login')
                for label in descendants(login):
                    if isinstance(label, app.ttk.Label) and str(label.cget('foreground')) == '#B42318':
                        login.setvar(str(label.cget('textvariable')), '网络连接暂时不可用，请稍后重试。' * 5)
                root.update()
                surface(login, login, 'login-long-error')
                login.destroy()
                gui.authorization_window = None
                form_root = app.tk.Toplevel(root)
                with patch.object(setup, 'load_connection', return_value={}), patch.object(setup, 'BitBrowserTokenStore') as store:
                    store.return_value.get_token.return_value = ''
                    form = setup.BitBrowserSetupWindow(form_root)
                form_root.geometry('720x430+20+20')
                root.update()
                surface(form_root, form_root, 'local-api-form')
                form_root.destroy()
    finally:
        root.destroy()
    return {'scaling': scaling, 'checked_count': len(checked), 'checked': checked, 'issues': issues}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', required=True)
    parser.add_argument('--scaling', nargs='+', type=float, default=[1.333, 1.667, 2.0])
    args = parser.parse_args()
    output = Path(args.output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    failed = False
    for scaling in args.scaling:
        report = audit_scale(scaling)
        (output / f'ui-buttons-{scaling:.3f}.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({key: report[key] for key in ('scaling', 'checked_count', 'issues')}, ensure_ascii=False), flush=True)
        failed |= bool(report['issues'])
    raise SystemExit(int(failed))


if __name__ == '__main__':
    main()
