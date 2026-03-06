#!/usr/bin/env python3
"""
DXF → GeoPackage 変換ツール — デスクトップアプリ
ネイティブウィンドウでFlaskアプリを表示する。
"""
import sys
import os
import shutil
import threading
import time

# PyInstaller bundled app: set correct base path
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
    # templates/static を PyInstaller バンドルから参照
    os.environ['FLASK_APP_BASE'] = BASE_DIR
    # GDAL/PROJ/Fiona データパスを設定（PyInstallerバンドル内）
    for subdir in ('fiona/gdal_data', 'gdal_data', 'gdal/data', 'share/gdal'):
        p = os.path.join(BASE_DIR, subdir)
        if os.path.isdir(p):
            os.environ['GDAL_DATA'] = p
            break
    for subdir in ('pyproj/proj_dir/share/proj', 'proj_data', 'share/proj'):
        p = os.path.join(BASE_DIR, subdir)
        if os.path.isdir(p):
            os.environ['PROJ_LIB'] = p
            break
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Ensure the working directory for temp files exists
WORK_DIR = os.path.join(os.path.expanduser('~'), '.dxf-to-gpkg-tmp')
os.makedirs(WORK_DIR, exist_ok=True)
os.environ['DXF_WORK_DIR'] = WORK_DIR

sys.path.insert(0, BASE_DIR)

PORT = 5050

# デフォルトの保存先
DOWNLOADS_DIR = os.path.join(os.path.expanduser('~'), 'Downloads')


class JsApi:
    """pywebview JavaScript API ブリッジ
    フロントエンドから window.pywebview.api.save_file(...) で呼び出される。
    """

    def save_file(self, job_id, filename):
        """ネイティブ保存ダイアログでGPKGファイルを保存する。

        Returns:
            dict: {'path': 保存先パス} or {'path': ..., 'default': True}（ダウンロードフォルダ）
                  or {'error': エラーメッセージ}
        """
        import webview

        # ソースファイルの確認
        src = os.path.join(WORK_DIR, job_id, filename)
        if not os.path.isfile(src):
            return {'error': 'ファイルが見つかりません'}

        try:
            # ネイティブ保存ダイアログを表示
            window = webview.windows[0]
            result = window.create_file_dialog(
                webview.SAVE_DIALOG,
                directory=DOWNLOADS_DIR,
                save_filename=filename,
            )

            if result:
                # ユーザーが保存先を選択した
                dest = result if isinstance(result, str) else result[0]
                shutil.copy2(src, dest)
                return {'path': dest}
            else:
                # キャンセル → ダウンロードフォルダにフォールバック
                dest = os.path.join(DOWNLOADS_DIR, filename)
                # 同名ファイルが存在する場合はリネーム
                if os.path.exists(dest):
                    base, ext = os.path.splitext(filename)
                    i = 1
                    while os.path.exists(dest):
                        dest = os.path.join(DOWNLOADS_DIR, f'{base}_{i}{ext}')
                        i += 1
                shutil.copy2(src, dest)
                return {'path': dest, 'default': True}

        except Exception as e:
            # ダイアログ表示失敗時もダウンロードフォルダにフォールバック
            try:
                dest = os.path.join(DOWNLOADS_DIR, filename)
                if os.path.exists(dest):
                    base, ext = os.path.splitext(filename)
                    i = 1
                    while os.path.exists(dest):
                        dest = os.path.join(DOWNLOADS_DIR, f'{base}_{i}{ext}')
                        i += 1
                shutil.copy2(src, dest)
                return {'path': dest, 'default': True}
            except Exception as e2:
                return {'error': str(e2)}


def start_server():
    """Start Flask server in background thread."""
    from web_app import app
    app.run(host='127.0.0.1', port=PORT, debug=False, use_reloader=False)


def wait_for_server():
    """Wait until Flask server is ready."""
    import urllib.request
    for _ in range(60):
        time.sleep(0.5)
        try:
            urllib.request.urlopen(f'http://localhost:{PORT}/')
            return True
        except Exception:
            continue
    return False


def main():
    print('=' * 50)
    print('  DXF → GeoPackage 変換ツール')
    print(f'  http://localhost:{PORT}')
    print('=' * 50)

    # Flask サーバーをバックグラウンドで起動
    server_thread = threading.Thread(target=start_server, daemon=True)
    server_thread.start()

    # サーバー起動待ち
    if not wait_for_server():
        print('サーバーの起動に失敗しました')
        sys.exit(1)

    try:
        import webview
        api = JsApi()
        # ネイティブウィンドウで表示（JS APIブリッジ付き）
        webview.create_window(
            'DXF → GeoPackage Converter',
            f'http://localhost:{PORT}',
            width=820,
            height=900,
            resizable=True,
            min_size=(600, 500),
            js_api=api,
        )
        webview.start()
    except ImportError:
        # pywebview が無い場合はブラウザにフォールバック
        import webbrowser
        print('  ブラウザで開きます...')
        webbrowser.open(f'http://localhost:{PORT}')
        # サーバーを維持
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            pass


if __name__ == '__main__':
    main()
