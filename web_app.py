#!/usr/bin/env python3
"""
DXF → GeoPackage 変換ツール Web UI
====================================
Flask + SSE によるリアルタイム変換進捗表示
localhost:5050 で駆動
"""

import os
import sys
import uuid
import time
import json
import shutil
import threading
from queue import Queue, Empty
from pathlib import Path
from datetime import datetime, timedelta

from flask import (
    Flask, render_template, request, jsonify,
    send_file, Response, stream_with_context
)

# converter.py を同ディレクトリからインポート
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from converter import (
    DxfConverter, save_to_geopackage,
    analyze_dxf_coordinates,
    JGD2011_EPSG, JGD2000_EPSG, ZONE_DESCRIPTIONS,
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

WORK_DIR = Path(__file__).parent / '.web_tmp'
WORK_DIR.mkdir(exist_ok=True)

# ジョブ管理: job_id -> {queue, status, result_files, ...}
jobs: dict[str, dict] = {}


def _cleanup_old_jobs(max_age_hours: int = 1):
    """古い一時ファイルを削除"""
    now = datetime.now()
    for d in WORK_DIR.iterdir():
        if d.is_dir():
            age = now - datetime.fromtimestamp(d.stat().st_mtime)
            if age > timedelta(hours=max_age_hours):
                shutil.rmtree(d, ignore_errors=True)
                jobs.pop(d.name, None)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload():
    """ファイルアップロード + 座標系分析"""
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルが選択されていません'}), 400

    f = request.files['file']
    if not f.filename:
        return jsonify({'error': 'ファイル名が空です'}), 400

    ext = Path(f.filename).suffix.lower()
    if ext not in ('.dxf',):
        return jsonify({'error': f'非対応の拡張子: {ext}（.dxf のみ対応）'}), 400

    job_id = str(uuid.uuid4())[:8]
    job_dir = WORK_DIR / job_id
    job_dir.mkdir(exist_ok=True)

    filepath = job_dir / f.filename
    f.save(str(filepath))

    # 座標系分析
    analysis = {}
    try:
        raw = analyze_dxf_coordinates(str(filepath))
        # JSON非対応型を変換
        analysis = {}
        for k, v in raw.items():
            if isinstance(v, set):
                analysis[k] = sorted(v)
            elif isinstance(v, list) and v and isinstance(v[0], tuple):
                analysis[k] = [list(t) for t in v]
            elif isinstance(v, float) and (v != v):  # NaN check
                analysis[k] = None
            else:
                analysis[k] = v
    except Exception as e:
        analysis = {'error': str(e)}

    # coord_range を整理してフロントエンドに渡す
    coord_range = None
    if analysis.get('x_min') is not None:
        coord_range = {
            'x_min': analysis['x_min'], 'x_max': analysis['x_max'],
            'y_min': analysis['y_min'], 'y_max': analysis['y_max'],
        }

    jobs[job_id] = {
        'filepath': str(filepath),
        'filename': f.filename,
        'analysis': analysis,
        'status': 'uploaded',
        'queue': None,
        'result_files': [],
    }

    return jsonify({
        'job_id': job_id,
        'filename': f.filename,
        'filesize': filepath.stat().st_size,
        'analysis': analysis,
        'coord_range': coord_range,
        'detected_zone': analysis.get('suggested_zone'),
        'zones': {str(k): v for k, v in ZONE_DESCRIPTIONS.items()},
    })


@app.route('/api/convert/<job_id>', methods=['POST'])
def convert(job_id):
    """変換開始（バックグラウンドスレッド）"""
    if job_id not in jobs:
        return jsonify({'error': 'ジョブが見つかりません'}), 404

    job = jobs[job_id]
    if job['status'] == 'converting':
        return jsonify({'error': '変換中です'}), 409

    params = request.get_json(silent=True) or {}
    scale = int(params.get('scale', 0))
    datum = params.get('datum', 'JGD2011')
    zone = int(params.get('zone', 9))
    output_crs = params.get('output_crs', 'source')
    custom_epsg = int(params.get('custom_epsg', 0))
    quality = int(params.get('quality', 100))
    auto_georef = bool(params.get('auto_georef', True))
    split_by_layer = bool(params.get('split_by_layer', True))

    # EPSG決定
    epsg_map = JGD2011_EPSG if datum == 'JGD2011' else JGD2000_EPSG
    source_epsg = epsg_map.get(zone, 6677)

    target_epsg = None
    if output_crs == 'wgs84':
        target_epsg = 4326
    elif output_crs == 'custom' and custom_epsg > 0:
        target_epsg = custom_epsg

    q = Queue()
    job['queue'] = q
    job['status'] = 'converting'
    job['result_files'] = []

    def worker():
        filepath = job['filepath']
        job_dir = Path(filepath).parent
        basename = Path(filepath).stem

        try:
            q.put({'type': 'log', 'level': 'info',
                   'msg': f'変換開始: {job["filename"]}'})
            q.put({'type': 'log', 'level': 'info',
                   'msg': f'縮尺: 1:{scale}, 系番号: {zone}, '
                          f'測地系: {datum}'})
            q.put({'type': 'progress', 'value': 5})

            # DXF変換
            q.put({'type': 'log', 'level': 'info',
                   'msg': 'DXF読込・変換中...'})
            conv = DxfConverter()
            result = conv.convert(
                filepath,
                scale_denominator=scale,
                quality=quality,
                auto_georef=auto_georef,
            )
            q.put({'type': 'progress', 'value': 50})

            # 警告表示
            for w in conv.warnings:
                q.put({'type': 'log', 'level': 'warning', 'msg': w})

            # フィーチャ数集計
            total = sum(len(gdf) for gdf in result.values() if not gdf.empty)
            layers = sum(1 for gdf in result.values() if not gdf.empty)
            q.put({'type': 'log', 'level': 'info',
                   'msg': f'変換完了: {total:,} フィーチャ, '
                          f'{layers} ジオメトリタイプ'})
            q.put({'type': 'progress', 'value': 70})

            # GeoPackage保存
            output_name = f'{basename}.gpkg'
            output_path = str(job_dir / output_name)
            q.put({'type': 'log', 'level': 'info',
                   'msg': 'GeoPackage保存中...'})

            success, msgs = save_to_geopackage(
                result, output_path,
                source_crs_epsg=source_epsg,
                target_crs_epsg=target_epsg,
                split_by_layer=split_by_layer,
            )
            q.put({'type': 'progress', 'value': 95})

            for m in msgs:
                lvl = 'warning' if 'エラー' in m or '警告' in m else 'info'
                q.put({'type': 'log', 'level': lvl, 'msg': m})

            if success:
                job['result_files'].append(output_name)
                fsize = os.path.getsize(output_path)
                q.put({'type': 'log', 'level': 'info',
                       'msg': f'出力: {output_name} '
                              f'({fsize / 1024 / 1024:.1f} MB)'})
                q.put({'type': 'progress', 'value': 100})
                q.put({'type': 'complete', 'files': job['result_files']})
                job['status'] = 'complete'
            else:
                q.put({'type': 'log', 'level': 'error',
                       'msg': '保存に失敗しました'})
                q.put({'type': 'error', 'msg': '保存に失敗しました'})
                job['status'] = 'error'

        except Exception as e:
            q.put({'type': 'log', 'level': 'error', 'msg': str(e)})
            q.put({'type': 'error', 'msg': str(e)})
            job['status'] = 'error'

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    return jsonify({'status': 'started'})


@app.route('/api/progress/<job_id>')
def progress(job_id):
    """SSE進捗ストリーム"""
    if job_id not in jobs:
        return jsonify({'error': 'ジョブが見つかりません'}), 404

    job = jobs[job_id]
    q = job.get('queue')
    if q is None:
        return jsonify({'error': 'キューがありません'}), 400

    def generate():
        while True:
            try:
                msg = q.get(timeout=30)
                yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
                if msg.get('type') in ('complete', 'error'):
                    break
            except Empty:
                # ハートビート
                yield f"data: {json.dumps({'type': 'heartbeat'})}\n\n"

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no',
        }
    )


@app.route('/api/download/<job_id>/<filename>')
def download(job_id, filename):
    """GPKG ダウンロード"""
    if job_id not in jobs:
        return jsonify({'error': 'ジョブが見つかりません'}), 404

    filepath = WORK_DIR / job_id / filename
    if not filepath.exists():
        return jsonify({'error': 'ファイルが見つかりません'}), 404

    return send_file(
        str(filepath),
        as_attachment=True,
        download_name=filename,
    )


@app.route('/api/cleanup/<job_id>', methods=['POST'])
def cleanup(job_id):
    """一時ファイル削除"""
    job_dir = WORK_DIR / job_id
    if job_dir.exists():
        shutil.rmtree(job_dir, ignore_errors=True)
    jobs.pop(job_id, None)
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    _cleanup_old_jobs()
    print("=" * 50)
    print("  DXF → GeoPackage 変換ツール")
    print("  http://localhost:5050")
    print("=" * 50)
    app.run(host='127.0.0.1', port=5050, debug=False)
