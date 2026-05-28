import multiprocessing
import streamlit.web.cli as stcli
import sys
import webview
import time
import os
# класс-мост
class WebviewAPI:
    def __init__(self):
        self._window = None

    def set_window(self, window):
        self._window = window

    def select_folder(self):
        if self._window:
            result = self._window.create_file_dialog(webview.FOLDER_DIALOG)
            if result:
                return result[0] if isinstance(result, (list, tuple)) else result
        return None

def run_streamlit():
    sys.argv = [
        "streamlit", "run", "app.py",
        "--global.developmentMode=false",
        "--server.headless=true",
        "--server.port=8501"
    ]
    stcli.main()

if __name__ == "__main__":
    p = multiprocessing.Process(target=run_streamlit)
    p.daemon = True
    p.start()

    time.sleep(1.5)

    api = WebviewAPI()

    window = webview.create_window(
        title="Программный комплекс РАО/ОЯТ",
        url="http://localhost:8501",
        width=1440,
        height=900,
        resizable=True,
        js_api=api
    )

    api.set_window(window)
    webview.start()

    p.terminate()
