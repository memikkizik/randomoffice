"""Top-level standalone desktop launcher for RandomOffice PyInstaller build."""

from pathlib import Path
import sys

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

if getattr(sys, "frozen", False):
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass and str(meipass) not in sys.path:
        sys.path.insert(0, str(meipass))

if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == '--self-test':
        import json
        import traceback
        import tempfile
        import time
        try:
            from randomoffice.desktop import locate_static_html, NativeDesktopBridge
            from randomoffice.ui import render_office_ui
            from randomoffice.web_office import export_office_file
            import webview
            import clr
            page = render_office_ui(locate_static_html())
            assert 'data-randomoffice-editor' in page
            for kind in ('word', 'excel', 'powerpoint'):
                assert len(export_office_file(kind, {})) > 1000
            test_directory = tempfile.TemporaryDirectory(prefix='randomoffice-startup-')
            startup_path=Path(test_directory.name) / 'Startup document.docx'
            startup_path.write_bytes(export_office_file('word',{'html':'<p>RandomOffice startup document</p>'}))
            bridge = NativeDesktopBridge(startup_document=startup_path, data_directory=Path(test_directory.name)/'data')
            window = webview.create_window('RandomOffice startup verification', html=page,
                                           hidden=True, js_api=bridge)
            bridge.set_window(window)
            ui_status = {}
            def check_loaded_ui():
                try:
                    bridge.notify_editor_ready()
                    deadline=time.monotonic()+30
                    ready=False
                    while time.monotonic()<deadline:
                        ready=window.evaluate_js("typeof nativeWorkspaceReady!=='undefined' && nativeWorkspaceReady && document.getElementById('word-editor').innerText.includes('RandomOffice startup document')")
                        if ready: break
                        time.sleep(.1)
                    if not ready: raise RuntimeError('Startup file was not delivered to the native editor')
                    window.evaluate_js("document.getElementById('word-editor').innerHTML='<p>Native draft verification</p>';persistState();")
                    while time.monotonic()<deadline:
                        stored=bridge.read_workspace_state().get('state',{})
                        if 'Native draft verification' in stored.get('wordHtml',''):
                            ui_status['ready']=True
                            break
                        time.sleep(.1)
                    if not ui_status.get('ready'): raise RuntimeError('Native draft was not persisted')
                except Exception:
                    ui_status['error'] = traceback.format_exc()
                finally:
                    window.destroy()
            window.events.loaded += check_loaded_ui
            webview.start(gui='edgechromium', debug=False)
            test_directory.cleanup()
            if not ui_status.get('ready'):
                raise RuntimeError('Native WebView UI failed: ' + str(ui_status))
            result = {'ok': True, 'frozen': bool(getattr(sys, 'frozen', False)),
                      'checks': ['python-dll', 'desktop-imports', 'dotnet', 'shared-ui', 'office-export', 'native-webview-ui','native-startup-document','native-draft-save']}
        except Exception:
            result = {'ok': False, 'error': traceback.format_exc()}
        Path(sys.argv[2]).write_text(json.dumps(result, indent=2), encoding='utf-8')
        raise SystemExit(0 if result['ok'] else 1)
    else:
        import argparse
        from randomoffice.desktop import run_desktop_app
        parser = argparse.ArgumentParser(description="RandomOffice masaüstü uygulaması")
        parser.add_argument("document", nargs="?", type=Path, help="Açılacak DOCX, XLSX veya PPTX belgesi")
        arguments = parser.parse_args()
        run_desktop_app(startup_document=arguments.document)

