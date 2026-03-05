#!/usr/bin/env python3
"""
SFC/DXF → GeoPackage Converter
================================
SFC（SXFフィーチャコメントファイル）及びDXFファイルを
QGISで快適に利用できるGeoPackageに変換するデスクトップアプリ。

Mac / Windows クロスプラットフォーム対応。
依存: Python 3.10+, ezdxf, geopandas, fiona, pyproj, shapely
"""

import os
import re
import sys
import math
import json
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

import ezdxf
import geopandas as gpd
import fiona
from shapely.geometry import Point, LineString, Polygon, MultiLineString, MultiPoint
from shapely.ops import unary_union
from pyproj import CRS, Transformer
import numpy as np


# ============================================================
# 平面直角座標系 (Japan Plane Rectangular CS) 定義
# ============================================================
JGD2011_EPSG = {
    1: 6669, 2: 6670, 3: 6671, 4: 6672, 5: 6673,
    6: 6674, 7: 6675, 8: 6676, 9: 6677, 10: 6678,
    11: 6679, 12: 6680, 13: 6681, 14: 6682, 15: 6683,
    16: 6684, 17: 6685, 18: 6686, 19: 6687,
}

JGD2000_EPSG = {
    1: 2443, 2: 2444, 3: 2445, 4: 2446, 5: 2447,
    6: 2448, 7: 2449, 8: 2450, 9: 2451, 10: 2452,
    11: 2453, 12: 2454, 13: 2455, 14: 2456, 15: 2457,
    16: 2458, 17: 2459, 18: 2460, 19: 2461,
}

ZONE_DESCRIPTIONS = {
    1: "I系 (長崎県,鹿児島県の一部)",
    2: "II系 (福岡,佐賀,熊本,大分,宮崎,鹿児島)",
    3: "III系 (山口,島根,広島)",
    4: "IV系 (香川,愛媛,徳島,高知)",
    5: "V系 (兵庫,鳥取,岡山)",
    6: "VI系 (京都,大阪,福井,滋賀,三重,奈良,和歌山)",
    7: "VII系 (石川,富山,岐阜,愛知)",
    8: "VIII系 (新潟,長野,山梨,静岡)",
    9: "IX系 (東京都,福島,栃木,茨城,埼玉,千葉,群馬,神奈川)",
    10: "X系 (青森,秋田,山形,岩手,宮城)",
    11: "XI系 (北海道西部)",
    12: "XII系 (北海道中央部)",
    13: "XIII系 (北海道東部)",
    14: "XIV系 (東京都の一部 小笠原)",
    15: "XV系 (沖縄県那覇周辺)",
    16: "XVI系 (沖縄県先島)",
    17: "XVII系 (沖縄県大東)",
    18: "XVIII系 (東京都小笠原の一部)",
    19: "XIX系 (東京都南鳥島)",
}

# 各系の原点（緯度, 経度）- WGS84逆変換で系番号推定に使用
ZONE_ORIGINS_LATLON = {
    1: (33.0, 129.5),   2: (33.0, 131.0),   3: (36.0, 132.1667),
    4: (33.0, 133.5),   5: (36.0, 134.3333), 6: (36.0, 136.0),
    7: (36.0, 137.1667), 8: (36.0, 138.5),   9: (36.0, 139.8333),
    10: (40.0, 140.8333), 11: (44.0, 140.25), 12: (44.0, 142.25),
    13: (44.0, 144.25), 14: (26.0, 142.0),   15: (26.0, 127.5),
    16: (26.0, 124.0),  17: (26.0, 131.0),   18: (20.0, 136.0),
    19: (26.0, 154.0),
}


def analyze_dxf_coordinates(filepath: str) -> dict:
    """DXFファイルの座標範囲を解析し、座標系を推定する"""
    result = {
        'x_min': None, 'x_max': None,
        'y_min': None, 'y_max': None,
        'coord_type': '不明',
        'suggested_zone': None,
        'entity_count': 0,
        'layer_names': set(),
    }

    try:
        doc = ezdxf.readfile(filepath)
    except Exception as e:
        result['error'] = str(e)
        return result

    msp = doc.modelspace()
    xs, ys = [], []
    # TEXT/MTEXTからグリッドラベル座標を抽出
    real_xs = []  # X=... の実座標値
    real_ys = []  # Y=... の実座標値

    for entity in msp:
        result['entity_count'] += 1
        dxf = entity.dxf
        if hasattr(dxf, 'layer'):
            result['layer_names'].add(dxf.layer)

        try:
            if hasattr(dxf, 'start'):
                xs.append(dxf.start.x); ys.append(dxf.start.y)
            if hasattr(dxf, 'end'):
                xs.append(dxf.end.x); ys.append(dxf.end.y)
            if hasattr(dxf, 'insert'):
                xs.append(dxf.insert.x); ys.append(dxf.insert.y)
            if hasattr(dxf, 'center'):
                xs.append(dxf.center.x); ys.append(dxf.center.y)
            if entity.dxftype() in ('LWPOLYLINE',):
                for p in entity.get_points(format='xy'):
                    xs.append(p[0]); ys.append(p[1])
        except:
            pass

        # TEXT/MTEXTからグリッドラベル座標値を抽出
        try:
            if entity.dxftype() in ('TEXT', 'MTEXT'):
                t = (dxf.text if hasattr(dxf, 'text') else '').strip()
                # X=3100, X=-36000 形式
                m = re.match(r'^X=(-?\d+\.?\d*)$', t)
                if m:
                    real_xs.append(float(m.group(1)))
                    continue
                # 0013=X 形式（逆順）
                m = re.match(r'^(\d{4,})=X$', t)
                if m:
                    real_xs.append(float(m.group(1)[::-1]))
                    continue
                # Y=-10200, Y=12345 形式
                m = re.match(r'^Y=(-?\d+\.?\d*)$', t)
                if m:
                    real_ys.append(float(m.group(1)))
        except:
            pass

    if not xs:
        return result

    result['x_min'] = min(xs)
    result['x_max'] = max(xs)
    result['y_min'] = min(ys)
    result['y_max'] = max(ys)

    # グリッドラベルから実座標が取れた場合はそちらで系番号を判定
    if real_xs and real_ys:
        cx = (min(real_xs) + max(real_xs)) / 2
        cy = (min(real_ys) + max(real_ys)) / 2
        result['coord_type'] = '平面直角座標系（グリッドラベルから検出）'
        result['real_x_range'] = [min(real_xs), max(real_xs)]
        result['real_y_range'] = [min(real_ys), max(real_ys)]

        candidates = []
        for zone_num, epsg in JGD2011_EPSG.items():
            if zone_num >= 14:
                continue
            try:
                t = Transformer.from_crs(
                    CRS.from_epsg(epsg), CRS.from_epsg(4326),
                    always_xy=True
                )
                lon, lat = t.transform(cy, cx)
                if 122 < lon < 154 and 20 < lat < 46:
                    lat0, lon0 = ZONE_ORIGINS_LATLON[zone_num]
                    lon_diff = abs(lon - lon0)
                    if lon_diff < 2.0:
                        candidates.append((zone_num, lon_diff, lon, lat))
            except:
                pass

        candidates.sort(key=lambda x: x[1])
        result['zone_candidates'] = candidates[:3]

        if candidates:
            best = candidates[0]
            result['suggested_zone'] = best[0]
            desc_lines = ['平面直角座標系（グリッドラベルから検出）— 候補:']
            for z, ld, lon, lat in candidates[:3]:
                marker = '★' if z == best[0] else '  '
                desc_lines.append(
                    f'  {marker} {z}系（{ZONE_DESCRIPTIONS[z]}）'
                    f'→ 緯度{lat:.4f}° 経度{lon:.4f}°'
                )
            result['coord_type'] = '\n'.join(desc_lines)
        return result

    x_abs = max(abs(result['x_min']), abs(result['x_max']))
    y_abs = max(abs(result['y_min']), abs(result['y_max']))

    # エンティティ座標からの推定（フォールバック）
    if 120 < x_abs < 155 and 20 < y_abs < 50:
        result['coord_type'] = '緯度経度 (WGS84/JGD2011)'
    elif x_abs < 500 and y_abs < 500:
        result['coord_type'] = 'ローカル座標 (図面原点基準の可能性)'
    elif x_abs > 500 or y_abs > 500:
        result['coord_type'] = '平面直角座標系の可能性'

        cx = (result['x_min'] + result['x_max']) / 2
        cy = (result['y_min'] + result['y_max']) / 2

        candidates = []
        for zone_num, epsg in JGD2011_EPSG.items():
            if zone_num >= 14:
                continue
            try:
                t = Transformer.from_crs(
                    CRS.from_epsg(epsg), CRS.from_epsg(4326),
                    always_xy=True
                )
                lon, lat = t.transform(cx, cy)
                if 128 < lon < 146 and 26 < lat < 46:
                    lat0, lon0 = ZONE_ORIGINS_LATLON[zone_num]
                    lon_diff = abs(lon - lon0)
                    if lon_diff < 1.5:
                        candidates.append((zone_num, lon_diff, lon, lat))
            except:
                pass

        candidates.sort(key=lambda x: x[1])
        result['zone_candidates'] = candidates[:3]

        if candidates:
            best = candidates[0]
            result['suggested_zone'] = best[0]
            desc_lines = [
                f'平面直角座標系の可能性 — 候補:'
            ]
            for z, ld, lon, lat in candidates[:3]:
                marker = '★' if z == best[0] else '  '
                desc_lines.append(
                    f'  {marker} {z}系（{ZONE_DESCRIPTIONS[z]}）'
                    f'→ 緯度{lat:.4f}° 経度{lon:.4f}°'
                )
            result['coord_type'] = '\n'.join(desc_lines)
    else:
        result['coord_type'] = '不明'

    return result


def analyze_sfc_coordinates(filepath: str) -> dict:
    """SFCファイルの座標範囲を解析する"""
    result = {
        'x_min': None, 'x_max': None,
        'y_min': None, 'y_max': None,
        'coord_type': '不明',
        'suggested_zone': None,
        'entity_count': 0,
        'layer_names': set(),
    }

    parser = SfcParser()
    success = parser.parse_file(filepath)
    if not success:
        return result

    xs, ys = [], []
    for feat in parser.features:
        result['entity_count'] += 1
        if hasattr(feat, 'x') and hasattr(feat, 'y'):
            xs.append(feat.x); ys.append(feat.y)
        if hasattr(feat, 'x1'):
            xs.append(feat.x1); ys.append(feat.y1)
            xs.append(feat.x2); ys.append(feat.y2)
        if hasattr(feat, 'cx'):
            xs.append(feat.cx); ys.append(feat.cy)
        if hasattr(feat, 'points'):
            for px, py in feat.points:
                xs.append(px); ys.append(py)

    result['layer_names'] = set(l.name for l in parser.layers.values())

    if not xs:
        return result

    result['x_min'] = min(xs)
    result['x_max'] = max(xs)
    result['y_min'] = min(ys)
    result['y_max'] = max(ys)

    # DXFと同じ推定ロジックを適用
    dxf_result = analyze_dxf_coordinates.__code__  # ロジック共有のため
    x_abs = max(abs(result['x_min']), abs(result['x_max']))
    y_abs = max(abs(result['y_min']), abs(result['y_max']))

    if x_abs > 500 or y_abs > 500:
        result['coord_type'] = '平面直角座標系の可能性'
        cx = (result['x_min'] + result['x_max']) / 2
        cy = (result['y_min'] + result['y_max']) / 2
        candidates = []
        for zone_num, epsg in JGD2011_EPSG.items():
            if zone_num >= 14:
                continue
            try:
                t = Transformer.from_crs(CRS.from_epsg(epsg), CRS.from_epsg(4326), always_xy=True)
                lon, lat = t.transform(cx, cy)
                if 128 < lon < 146 and 26 < lat < 46:
                    lat0, lon0 = ZONE_ORIGINS_LATLON[zone_num]
                    lon_diff = abs(lon - lon0)
                    if lon_diff < 1.5:
                        candidates.append((zone_num, lon_diff, lon, lat))
            except:
                pass
        candidates.sort(key=lambda x: x[1])
        result['zone_candidates'] = candidates[:3]
        if candidates:
            best = candidates[0]
            result['suggested_zone'] = best[0]
            desc_lines = ['平面直角座標系の可能性 — 候補:']
            for z, ld, lon, lat in candidates[:3]:
                marker = '★' if z == best[0] else '  '
                desc_lines.append(
                    f'  {marker} {z}系（{ZONE_DESCRIPTIONS[z]}）→ 緯度{lat:.4f}° 経度{lon:.4f}°'
                )
            result['coord_type'] = '\n'.join(desc_lines)
    elif 120 < x_abs < 155 and 20 < y_abs < 50:
        result['coord_type'] = '緯度経度 (WGS84/JGD2011)'
    elif x_abs < 500 and y_abs < 500:
        result['coord_type'] = 'ローカル座標の可能性'

    return result


# ============================================================
# SFC Parser - SXFフィーチャコメントファイルパーサー
# ============================================================
@dataclass
class SfcColor:
    """色定義"""
    index: int
    r: int = 0
    g: int = 0
    b: int = 0
    name: str = ""


@dataclass
class SfcLayer:
    """レイヤ定義"""
    index: int
    name: str = ""
    visible: bool = True


@dataclass
class SfcFeature:
    """SXFフィーチャ基底クラス"""
    feature_type: str
    layer_index: int = 0
    color_index: int = 0
    line_type_index: int = 0
    line_width_index: int = 0


@dataclass
class SfcPoint(SfcFeature):
    x: float = 0.0
    y: float = 0.0
    marker_code: int = 0
    rotation: float = 0.0
    scale: float = 1.0


@dataclass
class SfcLine(SfcFeature):
    x1: float = 0.0
    y1: float = 0.0
    x2: float = 0.0
    y2: float = 0.0


@dataclass
class SfcPolyline(SfcFeature):
    points: list = field(default_factory=list)  # [(x, y), ...]
    closed: bool = False


@dataclass
class SfcCircle(SfcFeature):
    cx: float = 0.0
    cy: float = 0.0
    radius: float = 0.0


@dataclass
class SfcArc(SfcFeature):
    cx: float = 0.0
    cy: float = 0.0
    radius: float = 0.0
    start_angle: float = 0.0
    end_angle: float = 0.0
    direction: int = 0  # 0=CCW, 1=CW


@dataclass
class SfcEllipse(SfcFeature):
    cx: float = 0.0
    cy: float = 0.0
    radius_x: float = 0.0
    radius_y: float = 0.0
    rotation: float = 0.0
    start_angle: float = 0.0
    end_angle: float = 0.0
    direction: int = 0


@dataclass
class SfcText(SfcFeature):
    x: float = 0.0
    y: float = 0.0
    text: str = ""
    height: float = 1.0
    width: float = 0.0
    spacing: float = 0.0
    rotation: float = 0.0
    slant: float = 0.0
    font_index: int = 0
    direct: int = 0  # 0=horizontal, 1=vertical
    b_pnt: int = 0  # base point


@dataclass
class SfcSpline(SfcFeature):
    points: list = field(default_factory=list)
    open_close: int = 0  # 0=open, 1=close


class SfcParser:
    """SFC (SXF Feature Comment) ファイルパーサー"""

    # SXF既定義色 (index: (R, G, B, name))
    PREDEFINED_COLORS = {
        1: (0, 0, 0, "Black"),
        2: (255, 0, 0, "Red"),
        3: (0, 255, 0, "Green"),
        4: (0, 0, 255, "Blue"),
        5: (255, 255, 0, "Yellow"),
        6: (255, 0, 255, "Magenta"),
        7: (0, 255, 255, "Cyan"),
        8: (255, 255, 255, "White"),
        9: (190, 75, 0, "Brown"),
        10: (255, 127, 0, "Orange"),
        11: (127, 255, 0, "Lime"),
        12: (0, 127, 0, "Dark Green"),
        13: (0, 127, 255, "Sky"),
        14: (0, 0, 127, "Dark Blue"),
        15: (127, 0, 255, "Violet"),
        16: (127, 127, 127, "Gray"),
    }

    def __init__(self):
        self.colors: dict[int, SfcColor] = {}
        self.layers: dict[int, SfcLayer] = {}
        self.features: list[SfcFeature] = []
        self.font_defs: dict[int, str] = {}
        self.line_type_defs: dict[int, str] = {}
        self.line_width_defs: dict[int, float] = {}
        self.paper_size = (420.0, 297.0)  # A3 default
        self.sheet_name = ""
        self.warnings: list[str] = []

        # 既定義色を初期化
        for idx, (r, g, b, name) in self.PREDEFINED_COLORS.items():
            self.colors[idx] = SfcColor(idx, r, g, b, name)

    def parse_file(self, filepath: str) -> bool:
        """SFCファイルをパースする"""
        try:
            # エンコーディング検出
            for enc in ['shift_jis', 'cp932', 'utf-8', 'euc-jp']:
                try:
                    with open(filepath, 'r', encoding=enc) as f:
                        content = f.read()
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            else:
                self.warnings.append("エンコーディングを検出できません")
                return False

            return self._parse_content(content)
        except Exception as e:
            self.warnings.append(f"ファイル読み込みエラー: {e}")
            traceback.print_exc()
            return False

    def _parse_content(self, content: str) -> bool:
        """SFCコンテンツをパース"""
        # コメント行を除去、空行スキップ
        lines = []
        for line in content.split('\n'):
            line = line.strip()
            if line and not line.startswith('/*'):
                lines.append(line)

        i = 0
        in_data_section = False

        while i < len(lines):
            line = lines[i]

            # フィーチャコメント行を識別
            # SFCのデータは "FEATURE_NAME(params...);" 形式

            # 複数行にまたがるフィーチャを結合
            if '(' in line and ';' not in line:
                # セミコロンが来るまで行を結合
                combined = line
                while i + 1 < len(lines) and ';' not in combined:
                    i += 1
                    combined += ' ' + lines[i]
                line = combined

            try:
                self._parse_line(line)
            except Exception as e:
                self.warnings.append(f"行 {i+1} パースエラー: {str(e)[:80]}")

            i += 1

        return len(self.features) > 0 or len(self.layers) > 0

    def _parse_line(self, line: str):
        """個別行をパース"""
        # フィーチャ名(パラメータ...); の形式を検出
        match = re.match(r'(\w+)\s*\((.*)\)\s*;', line, re.DOTALL)
        if not match:
            return

        feat_name = match.group(1).upper()
        params_str = match.group(2).strip()

        # パラメータを解析（カンマ区切り、文字列はシングルクォート）
        params = self._parse_params(params_str)

        # フィーチャタイプに応じた処理
        handler = getattr(self, f'_handle_{feat_name.lower()}', None)
        if handler:
            handler(params)
        else:
            # 未対応フィーチャの場合、汎用的に幾何情報の抽出を試みる
            self._handle_generic(feat_name, params)

    def _parse_params(self, params_str: str) -> list:
        """パラメータ文字列をリストに変換"""
        params = []
        current = ''
        in_string = False
        depth = 0

        for ch in params_str:
            if ch == "'" and not in_string:
                in_string = True
                current += ch
            elif ch == "'" and in_string:
                in_string = False
                current += ch
            elif ch == '(' and not in_string:
                depth += 1
                current += ch
            elif ch == ')' and not in_string:
                depth -= 1
                current += ch
            elif ch == ',' and not in_string and depth == 0:
                params.append(self._convert_param(current.strip()))
                current = ''
            else:
                current += ch

        if current.strip():
            params.append(self._convert_param(current.strip()))

        return params

    def _convert_param(self, val: str):
        """パラメータ値を適切な型に変換"""
        if not val:
            return ''
        if val.startswith("'") and val.endswith("'"):
            return val[1:-1]
        try:
            if '.' in val:
                return float(val)
            return int(val)
        except ValueError:
            return val

    # ----- フィーチャハンドラ -----

    def _handle_sfig_org(self, params):
        """用紙原点 (図面サイズ等)"""
        # SFIG_ORG(name, flag, x_len, y_len, ...)
        if len(params) >= 4:
            self.sheet_name = str(params[0]) if params[0] else ""
            try:
                self.paper_size = (float(params[2]), float(params[3]))
            except (ValueError, IndexError):
                pass

    def _handle_layer(self, params):
        """レイヤ定義"""
        if len(params) >= 3:
            idx = int(params[0])
            name = str(params[1])
            visible = int(params[2]) if len(params) > 2 else 1
            self.layers[idx] = SfcLayer(idx, name, visible != 0)

    def _handle_color_def(self, params):
        """色定義"""
        if len(params) >= 4:
            idx = int(params[0])
            r, g, b = int(params[1]), int(params[2]), int(params[3])
            self.colors[idx] = SfcColor(idx, r, g, b)

    def _handle_font_def(self, params):
        """フォント定義"""
        if len(params) >= 2:
            idx = int(params[0])
            name = str(params[1])
            self.font_defs[idx] = name

    def _handle_line_type_def(self, params):
        """線種定義"""
        if len(params) >= 2:
            idx = int(params[0])
            self.line_type_defs[idx] = str(params[1]) if len(params) > 1 else ""

    def _handle_width_def(self, params):
        """線幅定義"""
        if len(params) >= 2:
            idx = int(params[0])
            self.line_width_defs[idx] = float(params[1])

    def _handle_point_marker(self, params):
        """点マーカ"""
        # POINT_MARKER(layer, color, x, y, marker_code, rotation, scale)
        if len(params) >= 5:
            feat = SfcPoint(
                feature_type='point_marker',
                layer_index=int(params[0]),
                color_index=int(params[1]),
                x=float(params[2]),
                y=float(params[3]),
                marker_code=int(params[4]) if len(params) > 4 else 0,
                rotation=float(params[5]) if len(params) > 5 else 0.0,
                scale=float(params[6]) if len(params) > 6 else 1.0,
            )
            self.features.append(feat)

    def _handle_line(self, params):
        """線分"""
        # LINE(layer, color, type, width, x1, y1, x2, y2)
        if len(params) >= 8:
            feat = SfcLine(
                feature_type='line',
                layer_index=int(params[0]),
                color_index=int(params[1]),
                line_type_index=int(params[2]),
                line_width_index=int(params[3]),
                x1=float(params[4]),
                y1=float(params[5]),
                x2=float(params[6]),
                y2=float(params[7]),
            )
            self.features.append(feat)

    def _handle_polyline(self, params):
        """折線"""
        # POLYLINE(layer, color, type, width, n, x1, y1, ..., xn, yn)
        if len(params) >= 5:
            layer = int(params[0])
            color = int(params[1])
            ltype = int(params[2])
            lwidth = int(params[3])
            n = int(params[4])
            points = []
            for j in range(n):
                idx_x = 5 + j * 2
                idx_y = 6 + j * 2
                if idx_y < len(params):
                    points.append((float(params[idx_x]), float(params[idx_y])))
            feat = SfcPolyline(
                feature_type='polyline',
                layer_index=layer,
                color_index=color,
                line_type_index=ltype,
                line_width_index=lwidth,
                points=points,
                closed=False,
            )
            self.features.append(feat)

    def _handle_polygon(self, params):
        """多角形（閉じた折線）"""
        self._handle_polyline(params)
        if self.features and isinstance(self.features[-1], SfcPolyline):
            self.features[-1].closed = True

    def _handle_circle(self, params):
        """円"""
        # CIRCLE(layer, color, type, width, cx, cy, radius)
        if len(params) >= 7:
            feat = SfcCircle(
                feature_type='circle',
                layer_index=int(params[0]),
                color_index=int(params[1]),
                line_type_index=int(params[2]),
                line_width_index=int(params[3]),
                cx=float(params[4]),
                cy=float(params[5]),
                radius=float(params[6]),
            )
            self.features.append(feat)

    def _handle_arc(self, params):
        """円弧"""
        # ARC(layer, color, type, width, cx, cy, radius, start_angle, end_angle, direction)
        if len(params) >= 9:
            feat = SfcArc(
                feature_type='arc',
                layer_index=int(params[0]),
                color_index=int(params[1]),
                line_type_index=int(params[2]),
                line_width_index=int(params[3]),
                cx=float(params[4]),
                cy=float(params[5]),
                radius=float(params[6]),
                start_angle=float(params[7]),
                end_angle=float(params[8]),
                direction=int(params[9]) if len(params) > 9 else 0,
            )
            self.features.append(feat)

    def _handle_ellipse(self, params):
        """楕円"""
        if len(params) >= 10:
            feat = SfcEllipse(
                feature_type='ellipse',
                layer_index=int(params[0]),
                color_index=int(params[1]),
                line_type_index=int(params[2]),
                line_width_index=int(params[3]),
                cx=float(params[4]),
                cy=float(params[5]),
                radius_x=float(params[6]),
                radius_y=float(params[7]),
                rotation=float(params[8]) if len(params) > 8 else 0.0,
                start_angle=float(params[9]) if len(params) > 9 else 0.0,
                end_angle=float(params[10]) if len(params) > 10 else 360.0,
            )
            self.features.append(feat)

    def _handle_text(self, params):
        """文字列"""
        # TEXT(layer, color, font, str, x, y, height, width, spacing, rotation, slant, direct, b_pnt)
        if len(params) >= 6:
            feat = SfcText(
                feature_type='text',
                layer_index=int(params[0]),
                color_index=int(params[1]),
                font_index=int(params[2]) if len(params) > 2 else 0,
                text=str(params[3]) if len(params) > 3 else "",
                x=float(params[4]) if len(params) > 4 else 0.0,
                y=float(params[5]) if len(params) > 5 else 0.0,
                height=float(params[6]) if len(params) > 6 else 1.0,
                width=float(params[7]) if len(params) > 7 else 0.0,
                spacing=float(params[8]) if len(params) > 8 else 0.0,
                rotation=float(params[9]) if len(params) > 9 else 0.0,
                slant=float(params[10]) if len(params) > 10 else 0.0,
                direct=int(params[11]) if len(params) > 11 else 0,
                b_pnt=int(params[12]) if len(params) > 12 else 0,
            )
            self.features.append(feat)

    def _handle_spline(self, params):
        """スプライン"""
        if len(params) >= 5:
            layer = int(params[0])
            color = int(params[1])
            ltype = int(params[2])
            lwidth = int(params[3])
            n = int(params[4])
            points = []
            for j in range(n):
                idx_x = 5 + j * 2
                idx_y = 6 + j * 2
                if idx_y < len(params):
                    points.append((float(params[idx_x]), float(params[idx_y])))
            oc = int(params[5 + n * 2]) if 5 + n * 2 < len(params) else 0
            feat = SfcSpline(
                feature_type='spline',
                layer_index=layer,
                color_index=color,
                line_type_index=ltype,
                line_width_index=lwidth,
                points=points,
                open_close=oc,
            )
            self.features.append(feat)

    def _handle_generic(self, feat_name: str, params: list):
        """未対応フィーチャの汎用処理 - 座標っぽい数値ペアを抽出"""
        # SFIG_ORG等の非幾何フィーチャはスキップ
        skip_names = {
            'HEADER', 'ENDER', 'SFIG_ORG', 'SFIG_END',
            'SXF_HEADER', 'SXF_ENDER',
            'LAYER', 'COLOR_DEF', 'FONT_DEF',
            'LINE_TYPE_DEF', 'WIDTH_DEF',
            'COMPOSIT_CURVE', 'COMPOSIT_CURVE_DEF',
        }
        if feat_name in skip_names:
            return

        # 数値パラメータからポイントペアの抽出を試みる
        numbers = [p for p in params if isinstance(p, (int, float))]
        if len(numbers) >= 4:
            layer_idx = int(numbers[0]) if numbers[0] == int(numbers[0]) else 0
            color_idx = int(numbers[1]) if len(numbers) > 1 and numbers[1] == int(numbers[1]) else 0

    def _arc_to_points(self, cx, cy, radius, start_deg, end_deg, direction=0, segments=32):
        """円弧を点列に変換"""
        points = []
        start_rad = math.radians(start_deg)
        end_rad = math.radians(end_deg)

        if direction == 0:  # CCW
            if end_rad <= start_rad:
                end_rad += 2 * math.pi
        else:  # CW
            if start_rad <= end_rad:
                start_rad += 2 * math.pi

        for i in range(segments + 1):
            t = i / segments
            angle = start_rad + t * (end_rad - start_rad)
            x = cx + radius * math.cos(angle)
            y = cy + radius * math.sin(angle)
            points.append((x, y))
        return points

    def to_geodataframes(self) -> dict[str, gpd.GeoDataFrame]:
        """パース結果をジオメトリタイプ別のGeoDataFrameに変換"""
        point_records = []
        line_records = []
        polygon_records = []
        text_records = []

        for feat in self.features:
            layer_name = self.layers.get(feat.layer_index, SfcLayer(0, f"Layer_{feat.layer_index}")).name
            color = self.colors.get(feat.color_index, SfcColor(0, 0, 0, 0))
            color_hex = f"#{color.r:02x}{color.g:02x}{color.b:02x}"

            base_attrs = {
                'layer': layer_name,
                'layer_idx': feat.layer_index,
                'color': color_hex,
                'color_r': color.r,
                'color_g': color.g,
                'color_b': color.b,
                'line_type': feat.line_type_index,
                'line_width': feat.line_width_index,
                'feat_type': feat.feature_type,
            }

            if isinstance(feat, SfcPoint):
                geom = Point(feat.x, feat.y)
                rec = {**base_attrs, 'geometry': geom,
                       'marker_code': feat.marker_code,
                       'rotation': feat.rotation,
                       'scale': feat.scale}
                point_records.append(rec)

            elif isinstance(feat, SfcText):
                geom = Point(feat.x, feat.y)
                ogr_style = _build_ogr_label_style(
                    feat.text, '', feat.height, feat.rotation, color_hex)
                rec = {**base_attrs, 'geometry': geom,
                       'text': feat.text,
                       'text_height': feat.height,
                       'text_width': feat.width,
                       'rotation': feat.rotation,
                       'slant': feat.slant,
                       'spacing': feat.spacing,
                       'direct': feat.direct,
                       'b_pnt': feat.b_pnt,
                       'OGR_STYLE': ogr_style}
                text_records.append(rec)

            elif isinstance(feat, SfcLine):
                geom = LineString([(feat.x1, feat.y1), (feat.x2, feat.y2)])
                rec = {**base_attrs, 'geometry': geom}
                line_records.append(rec)

            elif isinstance(feat, SfcPolyline):
                if len(feat.points) >= 2:
                    if feat.closed and len(feat.points) >= 3:
                        pts = list(feat.points)
                        if pts[0] != pts[-1]:
                            pts.append(pts[0])
                        geom = Polygon(pts)
                        rec = {**base_attrs, 'geometry': geom}
                        polygon_records.append(rec)
                    else:
                        geom = LineString(feat.points)
                        rec = {**base_attrs, 'geometry': geom}
                        line_records.append(rec)

            elif isinstance(feat, SfcCircle):
                geom = Point(feat.cx, feat.cy)
                rec = {**base_attrs, 'geometry': geom,
                       'radius': feat.radius, 'feat_type': 'circle'}
                point_records.append(rec)

            elif isinstance(feat, SfcArc):
                pts = self._arc_to_points(
                    feat.cx, feat.cy, feat.radius,
                    feat.start_angle, feat.end_angle, feat.direction
                )
                if len(pts) >= 2:
                    geom = LineString(pts)
                    rec = {**base_attrs, 'geometry': geom,
                           'radius': feat.radius,
                           'start_angle': feat.start_angle,
                           'end_angle': feat.end_angle}
                    line_records.append(rec)

            elif isinstance(feat, SfcEllipse):
                # 楕円を点列で近似
                pts = []
                segments = 32
                for i in range(segments + 1):
                    angle = 2 * math.pi * i / segments
                    x = feat.cx + feat.radius_x * math.cos(angle) * math.cos(math.radians(feat.rotation)) \
                        - feat.radius_y * math.sin(angle) * math.sin(math.radians(feat.rotation))
                    y = feat.cy + feat.radius_x * math.cos(angle) * math.sin(math.radians(feat.rotation)) \
                        + feat.radius_y * math.sin(angle) * math.cos(math.radians(feat.rotation))
                    pts.append((x, y))
                if len(pts) >= 4:
                    geom = Polygon(pts)
                    rec = {**base_attrs, 'geometry': geom}
                    polygon_records.append(rec)

            elif isinstance(feat, SfcSpline):
                if len(feat.points) >= 2:
                    geom = LineString(feat.points)
                    rec = {**base_attrs, 'geometry': geom}
                    line_records.append(rec)

        result = {}
        if point_records:
            result['points'] = gpd.GeoDataFrame(point_records, geometry='geometry')
        if line_records:
            result['lines'] = gpd.GeoDataFrame(line_records, geometry='geometry')
        if polygon_records:
            result['polygons'] = gpd.GeoDataFrame(polygon_records, geometry='geometry')
        if text_records:
            result['text'] = gpd.GeoDataFrame(text_records, geometry='geometry')

        return result


# ============================================================
# OGR Feature Style ヘルパー + DXF属性変換
# ============================================================

# AutoCAD Color Index (ACI) 全256色テーブル
# Index 1-9: 基本色, 10-249: HSL色環, 250-255: グレースケール
_ACI_COLORS_256 = [
    '#000000',  # 0: BYBLOCK
    '#ff0000',  # 1: Red
    '#ffff00',  # 2: Yellow
    '#00ff00',  # 3: Green
    '#00ffff',  # 4: Cyan
    '#0000ff',  # 5: Blue
    '#ff00ff',  # 6: Magenta
    '#000000',  # 7: White→Black (GIS白背景での視認性のため反転)
    '#808080',  # 8: Dark Gray
    '#c0c0c0',  # 9: Light Gray
    '#ff0000', '#ff7f7f', '#a50000', '#a55252', '#7f0000',  # 10-14
    '#7f3f3f', '#4c0000', '#4c2626', '#ff3f00', '#ff9f7f',  # 15-19
    '#a52900', '#a56752', '#7f1f00', '#7f4f3f', '#4c1300',  # 20-24
    '#4c2f26', '#ff7f00', '#ffbf7f', '#a55200', '#a57c52',  # 25-29
    '#7f3f00', '#7f5f3f', '#4c2600', '#4c3926', '#ffbf00',  # 30-34
    '#ffdf7f', '#a57c00', '#a59152', '#7f5f00', '#7f6f3f',  # 35-39
    '#4c3900', '#4c4226', '#ffff00', '#ffff7f', '#a5a500',  # 40-44
    '#a5a552', '#7f7f00', '#7f7f3f', '#4c4c00', '#4c4c26',  # 45-49
    '#bfff00', '#dfff7f', '#7ca500', '#91a552', '#5f7f00',  # 50-54
    '#6f7f3f', '#394c00', '#424c26', '#7fff00', '#bfff7f',  # 55-59
    '#52a500', '#7ca552', '#3f7f00', '#5f7f3f', '#264c00',  # 60-64
    '#394c26', '#3fff00', '#9fff7f', '#29a500', '#67a552',  # 65-69
    '#1f7f00', '#4f7f3f', '#134c00', '#2f4c26', '#00ff00',  # 70-74
    '#7fff7f', '#00a500', '#52a552', '#007f00', '#3f7f3f',  # 75-79
    '#004c00', '#264c26', '#00ff3f', '#7fff9f', '#00a529',  # 80-84
    '#52a567', '#007f1f', '#3f7f4f', '#004c13', '#264c2f',  # 85-89
    '#00ff7f', '#7fffbf', '#00a552', '#52a57c', '#007f3f',  # 90-94
    '#3f7f5f', '#004c26', '#264c39', '#00ffbf', '#7fffdf',  # 95-99
    '#00a57c', '#52a591', '#007f5f', '#3f7f6f', '#004c39',  # 100-104
    '#264c42', '#00ffff', '#7fffff', '#00a5a5', '#52a5a5',  # 105-109
    '#007f7f', '#3f7f7f', '#004c4c', '#264c4c', '#00bfff',  # 110-114
    '#7fdfff', '#007ca5', '#5291a5', '#005f7f', '#3f6f7f',  # 115-119
    '#00394c', '#26424c', '#007fff', '#7fbfff', '#0052a5',  # 120-124
    '#527ca5', '#003f7f', '#3f5f7f', '#00264c', '#26394c',  # 125-129
    '#003fff', '#7f9fff', '#0029a5', '#5267a5', '#001f7f',  # 130-134
    '#3f4f7f', '#00134c', '#262f4c', '#0000ff', '#7f7fff',  # 135-139
    '#0000a5', '#5252a5', '#00007f', '#3f3f7f', '#00004c',  # 140-144
    '#26264c', '#3f00ff', '#9f7fff', '#2900a5', '#6752a5',  # 145-149
    '#1f007f', '#4f3f7f', '#13004c', '#2f264c', '#7f00ff',  # 150-154
    '#bf7fff', '#5200a5', '#7c52a5', '#3f007f', '#5f3f7f',  # 155-159
    '#26004c', '#39264c', '#bf00ff', '#df7fff', '#7c00a5',  # 160-164
    '#9152a5', '#5f007f', '#6f3f7f', '#39004c', '#42264c',  # 165-169
    '#ff00ff', '#ff7fff', '#a500a5', '#a552a5', '#7f007f',  # 170-174
    '#7f3f7f', '#4c004c', '#4c264c', '#ff00bf', '#ff7fdf',  # 175-179
    '#a5007c', '#a55291', '#7f005f', '#7f3f6f', '#4c0039',  # 180-184
    '#4c2642', '#ff007f', '#ff7fbf', '#a50052', '#a5527c',  # 185-189
    '#7f003f', '#7f3f5f', '#4c0026', '#4c2639', '#ff003f',  # 190-194
    '#ff7f9f', '#a50029', '#a55267', '#7f001f', '#7f3f4f',  # 195-199
    '#4c0013', '#4c262f', '#333333', '#505050', '#696969',  # 200-204
    '#828282', '#bebebe', '#ffffff', '#000000', '#000000',  # 205-209
    '#000000', '#000000', '#000000', '#000000', '#000000',  # 210-214
    '#000000', '#000000', '#000000', '#000000', '#000000',  # 215-219
    '#000000', '#000000', '#000000', '#000000', '#000000',  # 220-224
    '#000000', '#000000', '#000000', '#000000', '#000000',  # 225-229
    '#000000', '#000000', '#000000', '#000000', '#000000',  # 230-234
    '#000000', '#000000', '#000000', '#000000', '#000000',  # 235-239
    '#000000', '#000000', '#000000', '#000000', '#000000',  # 240-244
    '#000000', '#000000', '#000000', '#000000', '#000000',  # 245-249
    '#333333', '#464646', '#585858', '#6b6b6b', '#808080',  # 250-254
    '#ebebeb',  # 255
]


def _aci_to_hex(color_idx: int) -> str:
    """AutoCADカラーインデックスをHEXカラーに変換"""
    if 0 <= color_idx <= 255:
        return _ACI_COLORS_256[color_idx]
    return '#000000'


def _resolve_entity_color(entity, layer_table=None) -> str:
    """エンティティの実効色をHEXで返す（true_color > ACI > BYLAYER）"""
    dxf = entity.dxf
    # true_color (24bit RGB) が設定されている場合
    if hasattr(dxf, 'true_color') and dxf.true_color is not None and dxf.true_color != 0:
        tc = dxf.true_color
        r = (tc >> 16) & 0xFF
        g = (tc >> 8) & 0xFF
        b = tc & 0xFF
        return f'#{r:02x}{g:02x}{b:02x}'
    # ACI color
    color_idx = dxf.color if hasattr(dxf, 'color') else 256
    if color_idx == 256 and layer_table:  # BYLAYER
        layer_name = dxf.layer if hasattr(dxf, 'layer') else '0'
        if layer_name in layer_table:
            color_idx = layer_table[layer_name].get('color', 7)
    elif color_idx == 0:  # BYBLOCK → 黒
        color_idx = 7
    return _aci_to_hex(color_idx)


def _dxf_lineweight_mm(lw_val: int) -> float:
    """DXF lineweight値（1/100mm）を実際のmm幅に変換"""
    if lw_val < 0:  # -1=BYLAYER, -2=BYBLOCK, -3=DEFAULT
        return 0.25  # デフォルト幅
    return lw_val / 100.0


def _dxf_linetype_to_qgis_dash(linetype_name: str, doc=None) -> str:
    """DXF線種名からQGIS用カスタムダッシュパターン文字列を生成

    Returns:
        ダッシュパターン文字列 (例: "5;2;1;2") またはソリッドなら空文字列
    """
    name = linetype_name.strip()

    # ソリッド線
    if name in ('実線', 'Continuous', 'CONTINUOUS', 'ByLayer', 'BYLAYER',
                'ByBlock', 'BYBLOCK', ''):
        return ''
    # no_name系はソリッド（descriptionが _____ のもの）
    if name.startswith('no_name'):
        return ''

    # ezdxfのlinetypeパターンから計算
    if doc:
        try:
            lt = doc.linetypes.get(name)
            if lt and hasattr(lt, 'pattern_tags'):
                lengths = lt.pattern_tags.lengths
                if lengths:
                    dash_parts = []
                    for val in lengths:
                        if val == 0:
                            dash_parts.append('0.3')  # ドット
                        else:
                            dash_parts.append(str(round(abs(val), 2)))
                    if dash_parts:
                        return ';'.join(dash_parts)
        except Exception:
            pass

    # 名前ベースのSXF標準線種マッピング（値は図面単位=メートル）
    # 破線系
    if name in ('破線', '跳び破線'):
        return '3;1.5'
    if '点線' in name:
        return '0.3;1.5'

    # 三点系（先にチェック — '二点'より前に'三点'をマッチ）
    if '三点' in name and '短' in name:
        return '3;1;0.3;1;0.3;1;0.3;1'
    if '三点' in name:
        return '6;1;0.3;1;0.3;1;0.3;1'

    # 二点系
    if '二点' in name and '短' in name:
        return '3;1;0.3;1;0.3;1'
    if '二点' in name:
        return '6;1;0.3;1;0.3;1'

    # 一点系
    if '一点' in name and '短' in name:
        return '3;1;0.3;1'
    if '一点' in name or '鎖線' in name:
        return '6;1;0.3;1'

    # 説明文ベースのパターン推定（descriptionがあれば）
    if doc:
        try:
            lt = doc.linetypes.get(name)
            if lt:
                desc = lt.dxf.description if hasattr(lt.dxf, 'description') else ''
                dot_count = desc.count('･')
                if dot_count >= 3:
                    return '6;1;0.3;1;0.3;1;0.3;1'  # 三点鎖線
                elif dot_count == 2:
                    return '6;1;0.3;1;0.3;1'  # 二点鎖線
                elif dot_count == 1:
                    return '6;1;0.3;1'  # 一点鎖線
                elif '_ _' in desc:
                    return '3;1.5'  # 破線
                elif '___' in desc:
                    return ''  # ソリッド
        except Exception:
            pass

    return '3;1.5'  # 不明な線種は破線扱い


def _build_ogr_label_style(
    text: str,
    font: str = '',
    size: float = 1.0,
    angle: float = 0.0,
    color_hex: str = '#000000',
    anchor: int = 7,
) -> str:
    """OGR Feature Style LABEL文字列を生成

    QGIS が自動的にテキストのフォント・サイズ・回転・色を認識して
    ラベルとして表示するための標準フォーマット。

    参考: https://gdal.org/user/ogr_feature_style.html

    Args:
        text: 表示テキスト
        font: フォント名（空の場合はMS Gothic）
        size: テキスト高さ（図面単位=ground units, "g"サフィックス付き）
        angle: 回転角度（度、反時計回り）
        color_hex: 色（#RRGGBB形式）
        anchor: アンカーポイント (1-12, 7=lower-left)
    """
    if not font:
        font = 'MS Gothic'
    # フォント名から拡張子(.ttf, .shx等)を除去
    font = re.sub(r'\.(ttf|otf|shx|fon)$', '', font, flags=re.IGNORECASE)
    # テキスト内のダブルクォートをエスケープ
    escaped_text = text.replace('"', '""')
    # 色の正規化
    c = color_hex.lower() if color_hex.startswith('#') else f'#{color_hex}'.lower()
    # 数値の精度を適切に丸める
    s_str = f'{round(size, 4)}' if size != 0 else '1'
    a_str = f'{round(angle, 4)}'
    return f'LABEL(f:"{font}",t:"{escaped_text}",s:{s_str}g,a:{a_str},c:{c},p:{anchor})'


# ============================================================
# DXF → GeoPackage 変換
# ============================================================
class DxfConverter:
    """DXFファイルをGeoDataFrameに変換"""

    def __init__(self):
        self.warnings: list[str] = []

    @staticmethod
    def _decode_dxf_unicode_escapes(text: str) -> str:
        """DXF独自の\\U+nnnnエスケープをUnicode文字にデコード"""
        import re
        def replace_match(m):
            try:
                return chr(int(m.group(1), 16))
            except (ValueError, OverflowError):
                return m.group(0)
        return re.sub(r'\\U\+([0-9A-Fa-f]{4})', replace_match, text)

    @staticmethod
    def _decode_dxf_special_chars(text: str) -> str:
        """DXF特殊文字（%%記法）をデコード"""
        replacements = {
            '%%d': '°',     # 度
            '%%D': '°',
            '%%p': '±',     # プラスマイナス
            '%%P': '±',
            '%%c': 'Ø',     # 直径記号
            '%%C': 'Ø',
            '%%u': '',       # 下線開始/終了（テキストのみ抽出なので除去）
            '%%U': '',
            '%%o': '',       # 上線開始/終了
            '%%O': '',
            '%%%': '%',      # %そのもの
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        # %%nnn形式（ASCII文字コード）
        import re
        def replace_ascii(m):
            try:
                return chr(int(m.group(1)))
            except (ValueError, OverflowError):
                return m.group(0)
        text = re.sub(r'%%(\d{3})', replace_ascii, text)
        return text

    @staticmethod
    def _clean_mtext(entity) -> str:
        """MTEXTエンティティからプレーンテキストを抽出"""
        if hasattr(entity, 'plain_text'):
            text = entity.plain_text()
        else:
            text = str(entity.dxf.text) if hasattr(entity.dxf, 'text') else ''
        # 改行文字を統一
        text = text.replace('\n', ' ').strip()
        return text

    @staticmethod
    def _clean_text(text: str) -> str:
        """TEXT/ATTRIBのテキストをクリーンアップ"""
        text = DxfConverter._decode_dxf_unicode_escapes(text)
        text = DxfConverter._decode_dxf_special_chars(text)
        return text.strip()

    @staticmethod
    def _detect_encoding(filepath: str) -> str | None:
        """DXFファイルのエンコーディングを検出"""
        # まずバイナリで読んでDWGCODEPAGEを探す
        try:
            with open(filepath, 'rb') as f:
                raw = f.read(8192)  # ヘッダ部分だけ読む

            # $DWGCODEPAGEを探す
            idx = raw.find(b'$DWGCODEPAGE')
            if idx >= 0:
                # 値は数行先にある
                after = raw[idx:idx+200]
                lines = after.split(b'\n')
                for i, line in enumerate(lines):
                    line = line.strip()
                    if line.startswith(b'ANSI_'):
                        codepage = line.decode('ascii', errors='ignore')
                        cp_num = codepage.replace('ANSI_', '')
                        if cp_num == '932':
                            return 'cp932'
                        elif cp_num == '936':
                            return 'gbk'
                        elif cp_num == '949':
                            return 'euc-kr'
                        elif cp_num == '950':
                            return 'big5'
                        elif cp_num == '1252':
                            return 'cp1252'
                        else:
                            try:
                                return f'cp{cp_num}'
                            except:
                                pass

            # DWGCODEPAGEがない場合、Shift-JISバイトパターンで判定
            # 0x81-0x9F, 0xE0-0xFC = Shift-JIS先行バイト
            sjis_score = 0
            for i in range(len(raw) - 1):
                b = raw[i]
                if (0x81 <= b <= 0x9F or 0xE0 <= b <= 0xFC):
                    b2 = raw[i + 1]
                    if (0x40 <= b2 <= 0x7E or 0x80 <= b2 <= 0xFC):
                        sjis_score += 1

            if sjis_score > 5:
                return 'cp932'

        except Exception:
            pass
        return None

    @staticmethod
    def _get_text_anchor_point(dxf):
        """DXF TEXTの配置設定から正しい挿入点・OGRアンカー・QGIS Hali/Valiを返す

        DXF TEXT は halign/valign でテキストの基準点を指定する。
        デフォルト(h=0, v=0)では dxf.insert が基準点。
        それ以外では dxf.align_point が実際の基準点になる。

        Returns:
            (x, y, ogr_anchor, hali, vali):
                x, y = 正しいワールド座標
                ogr_anchor = OGR Feature Style アンカー (1-12)
                hali = QGIS水平配置 ("Left"/"Center"/"Right")
                vali = QGIS垂直配置 ("Top"/"Half"/"Bottom"/"Base")
        """
        halign = dxf.halign if hasattr(dxf, 'halign') else 0
        valign = dxf.valign if hasattr(dxf, 'valign') else 0

        # デフォルト配置 (left-baseline) の場合は insert を使用
        # それ以外は align_point を使用
        if halign == 0 and valign == 0:
            pt = dxf.insert
        elif hasattr(dxf, 'align_point'):
            pt = dxf.align_point
        else:
            pt = dxf.insert

        x, y = pt.x, pt.y

        # OGR anchor mapping
        if halign == 4:  # Middle (centered both ways)
            anchor = 5
        elif halign in (3, 5):  # Aligned / Fit
            anchor = 10  # baseline-left
        else:
            h_col = min(halign, 2)  # 0=left, 1=center, 2=right
            # valign: 0=Baseline, 1=Bottom, 2=Middle, 3=Top
            # → OGR row: 0=Top, 1=Mid, 2=Bot, 3=Base
            v_row_map = {0: 3, 1: 2, 2: 1, 3: 0}
            v_row = v_row_map.get(valign, 3)
            anchor = v_row * 3 + h_col + 1

        # QGIS Hali / Vali 文字列
        hali_map = {0: 'Left', 1: 'Center', 2: 'Right',
                    3: 'Left', 4: 'Center', 5: 'Left'}
        vali_map = {0: 'Base', 1: 'Bottom', 2: 'Half', 3: 'Top'}
        hali = hali_map.get(halign, 'Left')
        vali = vali_map.get(valign, 'Base')

        return x, y, anchor, hali, vali

    @staticmethod
    def _get_mtext_anchor(dxf):
        """MTEXT の attachment_point → OGRアンカー + QGIS Hali/Vali

        Returns:
            (ogr_anchor, hali, vali)
        """
        ap = dxf.attachment_point if hasattr(dxf, 'attachment_point') else 1
        if not (1 <= ap <= 9):
            ap = 1

        hali_map = {1: 'Left', 2: 'Center', 3: 'Right',
                    4: 'Left', 5: 'Center', 6: 'Right',
                    7: 'Left', 8: 'Center', 9: 'Right'}
        vali_map = {1: 'Top', 2: 'Top', 3: 'Top',
                    4: 'Half', 5: 'Half', 6: 'Half',
                    7: 'Bottom', 8: 'Bottom', 9: 'Bottom'}

        return ap, hali_map.get(ap, 'Left'), vali_map.get(ap, 'Top')

    @staticmethod
    def _circle_segments(radius: float, quality: int) -> int:
        """半径と品質レベルに応じた円のセグメント数を返す

        Args:
            radius: 円の半径(DXF単位=mm)
            quality: 品質 100=フル, 50=中, 30=軽量
        """
        if quality >= 100:
            return 64
        elif quality >= 50:
            # 適応セグメント（半径依存）
            if radius <= 0.3:
                return 8
            elif radius <= 1.0:
                return 16
            elif radius <= 3.0:
                return 24
            else:
                return 32
        else:
            # 軽量: 最小限
            if radius <= 0.3:
                return 6
            elif radius <= 1.0:
                return 8
            elif radius <= 3.0:
                return 12
            else:
                return 16

    @staticmethod
    def _compute_affine_from_grid(text_gdf) -> Optional[tuple]:
        """テキストレイヤーから方眼線ラベルを検出し、GCPからアフィン変換行列を計算

        Returns:
            (a, b, tx, c, d, ty) アフィン変換パラメータ or None
            Real_X = a * dxf_x + b * dxf_y + tx
            Real_Y = c * dxf_x + d * dxf_y + ty
        """
        import re
        import numpy as np

        # X方眼線: "0013=X" → 反転 → X=3100
        x_lines = {}  # {real_X: [(dxf_x, dxf_y), ...]}
        # Y方眼線: "Y=-10200"
        y_lines = {}  # {real_Y: [(dxf_x, dxf_y), ...]}

        for _, row in text_gdf.iterrows():
            t = str(row.get('text', ''))
            geom = row.geometry
            if geom is None:
                continue

            # X方眼線パターン1: "0013=X" → 逆順 → X=3100（旧ソート順）
            m = re.match(r'^(\d{4,})=X$', t)
            if m:
                digits = m.group(1)
                real_x = float(digits[::-1])
                x_lines.setdefault(real_x, []).append((geom.x, geom.y))
                continue

            # X方眼線パターン2: "X=3100"（新ソート順 = 正しい読み順）
            m = re.match(r'^X=(-?\d+\.?\d*)$', t)
            if m:
                real_x = float(m.group(1))
                x_lines.setdefault(real_x, []).append((geom.x, geom.y))
                continue

            # Y方眼線パターン: "Y=-10200"
            m = re.match(r'^Y=(-?\d+\.?\d*)$', t)
            if m:
                real_y = float(m.group(1))
                y_lines.setdefault(real_y, []).append((geom.x, geom.y))

        if len(x_lines) < 2 or len(y_lines) < 2:
            return None

        def line_eq(p1, p2):
            """2点から ax + by = c"""
            a = -(p2[1] - p1[1])
            b = p2[0] - p1[0]
            c = a * p1[0] + b * p1[1]
            return (a, b, c)

        def intersect(l1, l2):
            a1, b1, c1 = l1
            a2, b2, c2 = l2
            det = a1 * b2 - a2 * b1
            if abs(det) < 1e-10:
                return None
            return ((c1 * b2 - c2 * b1) / det, (a1 * c2 - a2 * c1) / det)

        # 方眼線方程式
        x_eqs = {}
        for xv, pts in x_lines.items():
            if len(pts) >= 2:
                pts_sorted = sorted(pts, key=lambda p: p[1])
                x_eqs[xv] = line_eq(pts_sorted[0], pts_sorted[-1])

        y_eqs = {}
        for yv, pts in y_lines.items():
            if len(pts) >= 2:
                pts_sorted = sorted(pts, key=lambda p: p[0])
                y_eqs[yv] = line_eq(pts_sorted[0], pts_sorted[-1])

        # 交点 → GCP
        gcps = []
        for xv, xl in x_eqs.items():
            for yv, yl in y_eqs.items():
                pt = intersect(xl, yl)
                if pt:
                    gcps.append((pt[0], pt[1], xv, yv))

        if len(gcps) < 3:
            return None

        # 最小二乗法でアフィン変換
        src = np.array([[g[0], g[1], 1.0] for g in gcps])
        dst_x = np.array([g[2] for g in gcps])
        dst_y = np.array([g[3] for g in gcps])

        res_x, _, _, _ = np.linalg.lstsq(src, dst_x, rcond=None)
        res_y, _, _, _ = np.linalg.lstsq(src, dst_y, rcond=None)

        # (a, b, tx, c, d, ty)
        return (res_x[0], res_x[1], res_x[2],
                res_y[0], res_y[1], res_y[2])

    def convert(self, filepath: str, scale_denominator: int = 0,
                quality: int = 100,
                auto_georef: bool = True) -> dict[str, gpd.GeoDataFrame]:
        """DXFファイルを読み込み、ジオメトリタイプ別GeoDataFrameを返す

        Args:
            filepath: DXFファイルパス
            scale_denominator: 図面縮尺の分母（例: 300で1:300）
                設定時、text_height/dash_pattern/座標を実寸スケーリング
                0=スケーリングなし
            quality: 出力品質 (100=フル, 50=中品質, 30=軽量)
                CIRCLEのセグメント数を制御しファイルサイズに影響
        """
        # スケールファクター（mm→実寸m変換: 1:300なら 300/1000=0.3）
        scale_factor = scale_denominator / 1000.0 if scale_denominator > 0 else 0
        # エンコーディング検出
        detected_enc = self._detect_encoding(filepath)

        # まずezdxfの自動検出で試す
        doc = None
        try:
            doc = ezdxf.readfile(filepath)
        except Exception:
            pass

        # 自動検出で失敗 or 日本語DXFの疑いがある場合
        if doc is None or detected_enc == 'cp932':
            for enc in ['cp932', 'utf-8', 'cp1252']:
                try:
                    doc = ezdxf.readfile(filepath, encoding=enc)
                    self.warnings.append(f"エンコーディング: {enc} で読み込み")
                    break
                except Exception:
                    continue

        if doc is None:
            self.warnings.append("DXF読み込みエラー: 対応可能なエンコーディングが見つかりません")
            return {}

        # テキストスタイル情報を収集（DXFフォント名 → QGISフォントファミリー名）
        _FONT_FAMILY_MAP = {
            'msgothic.ttc': 'MS Gothic',
            'msgothic': 'MS Gothic',
            'msgothic.ttf': 'MS Gothic',
            'ms gothic': 'MS Gothic',
            'ＭＳ ゴシック': 'MS Gothic',
            'msゴシック': 'MS Gothic',
            'mspgoth.ttc': 'MS PGothic',
            'mspgothic': 'MS PGothic',
            'ms pgothic': 'MS PGothic',
            'ＭＳ Ｐゴシック': 'MS PGothic',
            'msmincho.ttc': 'MS Mincho',
            'msmincho': 'MS Mincho',
            'ms mincho': 'MS Mincho',
            'ＭＳ 明朝': 'MS Mincho',
            'mspmincho.ttc': 'MS PMincho',
            'ms pmincho': 'MS PMincho',
            'ＭＳ Ｐ明朝': 'MS PMincho',
            'yugothic': 'Yu Gothic',
            'yugothb.ttc': 'Yu Gothic',
            'yugothm.ttc': 'Yu Gothic',
            'yumincho': 'Yu Mincho',
            'arial.ttf': 'Arial',
            'arial': 'Arial',
            'times.ttf': 'Times New Roman',
            'times': 'Times New Roman',
            'cour.ttf': 'Courier New',
            'courier': 'Courier New',
        }
        def _map_font_family(dxf_font):
            """DXFフォント名をQGISフォントファミリー名にマッピング"""
            if not dxf_font:
                return 'MS Gothic'
            key = dxf_font.lower().strip()
            if key in _FONT_FAMILY_MAP:
                return _FONT_FAMILY_MAP[key]
            # 部分一致で検索
            for k, v in _FONT_FAMILY_MAP.items():
                if k in key or key in k:
                    return v
            # マッピングにない場合はそのまま返す（QGISが最善マッチを試みる）
            return dxf_font

        text_styles = {}  # style_name → (dxf_font_file, qgis_font_family)
        try:
            for style in doc.styles:
                name = style.dxf.name
                font = style.dxf.font if hasattr(style.dxf, 'font') else ''
                text_styles[name] = font
        except Exception:
            pass

        # レイヤーテーブルを収集（BYLAYER色の解決用）
        layer_table = {}
        try:
            for lyr in doc.layers:
                ld = lyr.dxf
                layer_table[ld.name] = {
                    'color': ld.color if hasattr(ld, 'color') else 7,
                    'linetype': ld.linetype if hasattr(ld, 'linetype') else '',
                    'lineweight': ld.lineweight if hasattr(ld, 'lineweight') else -3,
                }
        except Exception:
            pass

        msp = doc.modelspace()

        point_records = []
        line_records = []
        polygon_records = []
        text_records = []

        for entity in msp:
            dxf = entity.dxf
            layer = dxf.layer if hasattr(dxf, 'layer') else "0"
            color = dxf.color if hasattr(dxf, 'color') else 256

            # 図面属性を解決
            color_hex = _resolve_entity_color(entity, layer_table)
            linetype = dxf.linetype if hasattr(dxf, 'linetype') else ''
            lineweight = dxf.lineweight if hasattr(dxf, 'lineweight') else -3
            # BYLAYER解決
            if not linetype or linetype in ('ByLayer', 'BYLAYER'):
                linetype = layer_table.get(layer, {}).get('linetype', '実線')
            if lineweight < 0:
                lineweight = layer_table.get(layer, {}).get('lineweight', 25)
                if lineweight < 0:
                    lineweight = 25  # デフォルト

            lw_mm = _dxf_lineweight_mm(lineweight)
            dash_pattern = _dxf_linetype_to_qgis_dash(linetype, doc)

            # スケーリング適用（ダッシュパターン）
            if scale_factor > 0 and dash_pattern:
                dash_pattern = ';'.join(
                    str(round(float(v) * scale_factor, 4))
                    for v in dash_pattern.split(';')
                )

            # Qt pen style（QGIS outlineStyle用）
            pen_style = 'solid'
            lt_lower = linetype
            if '点線' in lt_lower:
                pen_style = 'dot'
            elif '二点' in lt_lower or '三点' in lt_lower:
                pen_style = 'dash dot dot'
            elif '一点' in lt_lower or '鎖線' in lt_lower:
                pen_style = 'dash dot'
            elif '破線' in lt_lower or '跳び' in lt_lower:
                pen_style = 'dash'
            elif dash_pattern:  # その他の非ソリッド
                pen_style = 'dash'

            base_attrs = {
                'layer': layer,
                'color_idx': color,
                'color_hex': color_hex,
                'linetype': linetype,
                'lineweight': lw_mm,
                'pen_style': pen_style,
                'dash_pattern': dash_pattern,
                'entity_type': entity.dxftype(),
            }

            try:
                etype = entity.dxftype()

                if etype == 'POINT':
                    pt = dxf.location
                    geom = Point(pt.x, pt.y)
                    point_records.append({**base_attrs, 'geometry': geom})

                elif etype == 'LINE':
                    p1, p2 = dxf.start, dxf.end
                    geom = LineString([(p1.x, p1.y), (p2.x, p2.y)])
                    line_records.append({**base_attrs, 'geometry': geom})

                elif etype in ('LWPOLYLINE', 'POLYLINE'):
                    if etype == 'LWPOLYLINE':
                        pts = [(p[0], p[1]) for p in entity.get_points(format='xy')]
                    else:
                        pts = [(v.dxf.location.x, v.dxf.location.y) for v in entity.vertices]

                    if len(pts) >= 2:
                        # LWPOLYLINE→closed, Polyline→is_closed
                        if hasattr(entity, 'closed'):
                            closed = entity.closed
                        elif hasattr(entity, 'is_closed'):
                            closed = entity.is_closed
                        else:
                            closed = bool(entity.dxf.flags & 1) if hasattr(entity.dxf, 'flags') else False
                        if closed and len(pts) >= 3:
                            if pts[0] != pts[-1]:
                                pts.append(pts[0])
                            geom = Polygon(pts)
                            polygon_records.append({**base_attrs, 'geometry': geom})
                        else:
                            geom = LineString(pts)
                            line_records.append({**base_attrs, 'geometry': geom})

                elif etype == 'CIRCLE':
                    cx, cy = dxf.center.x, dxf.center.y
                    r = dxf.radius
                    n_seg = self._circle_segments(r, quality)
                    circle_pts = []
                    for i in range(n_seg + 1):
                        a = 2 * math.pi * i / n_seg
                        circle_pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
                    geom = Polygon(circle_pts)
                    polygon_records.append({**base_attrs, 'geometry': geom,
                                            'radius': r})

                elif etype == 'ARC':
                    cx, cy = dxf.center.x, dxf.center.y
                    radius = dxf.radius
                    start_angle = math.radians(dxf.start_angle)
                    end_angle = math.radians(dxf.end_angle)
                    if end_angle <= start_angle:
                        end_angle += 2 * math.pi
                    # 角度に応じたセグメント数（最低16、最大64）
                    arc_span = end_angle - start_angle
                    n_seg = max(16, int(64 * arc_span / (2 * math.pi)))
                    pts = []
                    for i in range(n_seg + 1):
                        a = start_angle + arc_span * i / n_seg
                        pts.append((cx + radius * math.cos(a), cy + radius * math.sin(a)))
                    if len(pts) >= 2:
                        geom = LineString(pts)
                        line_records.append({**base_attrs, 'geometry': geom,
                                             'radius': radius})

                elif etype == 'ELLIPSE':
                    center = dxf.center
                    major_axis = dxf.major_axis
                    ratio = dxf.ratio
                    rx = math.sqrt(major_axis.x**2 + major_axis.y**2)
                    ry = rx * ratio
                    rot = math.atan2(major_axis.y, major_axis.x)
                    pts = []
                    segments = 32
                    for i in range(segments + 1):
                        a = 2 * math.pi * i / segments
                        x = center.x + rx * math.cos(a) * math.cos(rot) - ry * math.sin(a) * math.sin(rot)
                        y = center.y + rx * math.cos(a) * math.sin(rot) + ry * math.sin(a) * math.cos(rot)
                        pts.append((x, y))
                    if len(pts) >= 4:
                        geom = Polygon(pts)
                        polygon_records.append({**base_attrs, 'geometry': geom})

                elif etype in ('TEXT', 'MTEXT'):
                    if etype == 'TEXT':
                        # halign/valign に基づく正しい挿入点・アンカー・配置
                        tx, ty, anchor, hali, vali = self._get_text_anchor_point(dxf)
                        text = self._clean_text(dxf.text)
                        height = dxf.height if hasattr(dxf, 'height') else 1.0
                        rotation = dxf.rotation if hasattr(dxf, 'rotation') else 0.0
                        style = dxf.style if hasattr(dxf, 'style') else 'Standard'
                    else:  # MTEXT
                        tx, ty = dxf.insert.x, dxf.insert.y
                        anchor, hali, vali = self._get_mtext_anchor(dxf)
                        text = self._clean_mtext(entity)
                        text = self._decode_dxf_unicode_escapes(text)
                        height = dxf.char_height if hasattr(dxf, 'char_height') else 1.0
                        rotation = dxf.rotation if hasattr(dxf, 'rotation') else 0.0
                        style = dxf.style if hasattr(dxf, 'style') else 'Standard'

                    if text:  # 空テキストはスキップ
                        if scale_factor > 0:
                            height = height * scale_factor
                        font = _map_font_family(text_styles.get(style, ''))
                        color_hex = _aci_to_hex(color)
                        ogr_style = _build_ogr_label_style(
                            text, font, height, rotation, color_hex, anchor)
                        geom = Point(tx, ty)
                        rec = {**base_attrs, 'geometry': geom,
                               'text': text,
                               'text_height': height,
                               'rotation': rotation,
                               'font': font,
                               'style': style,
                               'color_hex': color_hex,
                               'char_spacing': 0.0,
                               'char_width_ratio': 0.0,
                               'hali': hali,
                               'vali': vali,
                               'OGR_STYLE': ogr_style}
                        text_records.append(rec)

                elif etype == 'INSERT':
                    # ブロック参照 — virtual_entities()で展開
                    block_name = dxf.name if hasattr(dxf, 'name') else ''

                    # ATTRIBテキストを抽出
                    if hasattr(entity, 'attribs'):
                        for attrib in entity.attribs:
                            att_text = self._clean_text(attrib.dxf.text)
                            if att_text:
                                att_x, att_y, att_anchor, att_hali, att_vali = self._get_text_anchor_point(attrib.dxf)
                                att_height = attrib.dxf.height if hasattr(attrib.dxf, 'height') else 1.0
                                if scale_factor > 0:
                                    att_height = att_height * scale_factor
                                att_rotation = attrib.dxf.rotation if hasattr(attrib.dxf, 'rotation') else 0.0
                                att_tag = attrib.dxf.tag if hasattr(attrib.dxf, 'tag') else ''
                                att_style = attrib.dxf.style if hasattr(attrib.dxf, 'style') else 'Standard'
                                font = _map_font_family(text_styles.get(att_style, ''))
                                att_color = attrib.dxf.color if hasattr(attrib.dxf, 'color') else color
                                att_color_hex = _aci_to_hex(att_color)
                                ogr_style = _build_ogr_label_style(
                                    att_text, font, att_height, att_rotation, att_color_hex, att_anchor)
                                geom = Point(att_x, att_y)
                                rec = {**base_attrs, 'geometry': geom,
                                       'entity_type': 'ATTRIB',
                                       'text': att_text,
                                       'text_height': att_height,
                                       'rotation': att_rotation,
                                       'font': font,
                                       'style': att_style,
                                       'attrib_tag': att_tag,
                                       'color_hex': att_color_hex,
                                       'char_spacing': 0.0,
                                       'char_width_ratio': 0.0,
                                       'hali': att_hali,
                                       'vali': att_vali,
                                       'OGR_STYLE': ogr_style}
                                text_records.append(rec)

                    # virtual_entities()でブロック内容をワールド座標で展開
                    try:
                        ve_texts = []  # (x, y, text, height, rotation, anchor, hali, vali, color_hex)
                        ve_height = 1.0
                        ve_style = 'Standard'
                        ve_anchor = 7  # default: bottom-left
                        ve_hali = 'Left'
                        ve_vali = 'Base'

                        for ve in entity.virtual_entities():
                            ve_type = ve.dxftype()
                            ve_dxf = ve.dxf
                            ve_layer = ve_dxf.layer if hasattr(ve_dxf, 'layer') else layer
                            ve_color = ve_dxf.color if hasattr(ve_dxf, 'color') else color
                            ve_color_hex = _resolve_entity_color(ve, layer_table)
                            ve_lt = ve_dxf.linetype if hasattr(ve_dxf, 'linetype') else linetype
                            ve_lw = ve_dxf.lineweight if hasattr(ve_dxf, 'lineweight') else lineweight
                            if not ve_lt or ve_lt in ('ByLayer', 'BYLAYER'):
                                ve_lt = layer_table.get(ve_layer, {}).get('linetype', '実線')
                            if ve_lw < 0:
                                ve_lw = layer_table.get(ve_layer, {}).get('lineweight', 25)
                                if ve_lw < 0:
                                    ve_lw = 25
                            ve_lw_mm = _dxf_lineweight_mm(ve_lw)
                            ve_dash = _dxf_linetype_to_qgis_dash(ve_lt, doc)
                            # スケーリング適用（ダッシュパターン）
                            if scale_factor > 0 and ve_dash:
                                ve_dash = ';'.join(
                                    str(round(float(v) * scale_factor, 4))
                                    for v in ve_dash.split(';')
                                )
                            # Qt pen style
                            ve_pen_style = 'solid'
                            if '点線' in ve_lt:
                                ve_pen_style = 'dot'
                            elif '二点' in ve_lt or '三点' in ve_lt:
                                ve_pen_style = 'dash dot dot'
                            elif '一点' in ve_lt or '鎖線' in ve_lt:
                                ve_pen_style = 'dash dot'
                            elif '破線' in ve_lt or '跳び' in ve_lt:
                                ve_pen_style = 'dash'
                            elif ve_dash:
                                ve_pen_style = 'dash'
                            ve_attrs = {
                                'layer': ve_layer,
                                'color_idx': ve_color,
                                'color_hex': ve_color_hex,
                                'linetype': ve_lt,
                                'lineweight': ve_lw_mm,
                                'pen_style': ve_pen_style,
                                'dash_pattern': ve_dash,
                                'entity_type': ve_type,
                            }

                            if ve_type == 'TEXT':
                                bt = self._clean_text(ve_dxf.text)
                                if bt and bt.strip():
                                    # halign/valign を考慮した正しい挿入点
                                    bx, by, b_anchor, b_hali, b_vali = self._get_text_anchor_point(ve_dxf)
                                    bh = ve_dxf.height if hasattr(ve_dxf, 'height') else 1.0
                                    br = ve_dxf.rotation if hasattr(ve_dxf, 'rotation') else 0.0
                                    ve_texts.append((bx, by, bt, bh, br, b_anchor, b_hali, b_vali, ve_color_hex))
                                    ve_height = bh
                                    ve_anchor = b_anchor
                                    ve_hali = b_hali
                                    ve_vali = b_vali
                                    if hasattr(ve_dxf, 'style'):
                                        ve_style = ve_dxf.style

                            elif ve_type == 'MTEXT':
                                bt = self._clean_mtext(ve)
                                bt = self._decode_dxf_unicode_escapes(bt)
                                if bt and bt.strip():
                                    bx = ve_dxf.insert.x
                                    by = ve_dxf.insert.y
                                    b_anchor, b_hali, b_vali = self._get_mtext_anchor(ve_dxf)
                                    bh = ve_dxf.char_height if hasattr(ve_dxf, 'char_height') else 1.0
                                    br = ve_dxf.rotation if hasattr(ve_dxf, 'rotation') else 0.0
                                    ve_texts.append((bx, by, bt, bh, br, b_anchor, b_hali, b_vali, ve_color_hex))
                                    ve_height = bh
                                    ve_anchor = b_anchor
                                    ve_hali = b_hali
                                    ve_vali = b_vali
                                    if hasattr(ve_dxf, 'style'):
                                        ve_style = ve_dxf.style

                            elif ve_type == 'LINE':
                                p1, p2 = ve_dxf.start, ve_dxf.end
                                geom = LineString([(p1.x, p1.y), (p2.x, p2.y)])
                                line_records.append({**ve_attrs, 'geometry': geom})

                            elif ve_type in ('LWPOLYLINE', 'POLYLINE'):
                                try:
                                    if ve_type == 'LWPOLYLINE':
                                        pts = [(p[0], p[1]) for p in ve.get_points(format='xy')]
                                    else:
                                        pts = [(v.dxf.location.x, v.dxf.location.y) for v in ve.vertices]
                                    if len(pts) >= 2:
                                        if hasattr(ve, 'closed'):
                                            closed = ve.closed
                                        elif hasattr(ve, 'is_closed'):
                                            closed = ve.is_closed
                                        else:
                                            closed = bool(ve.dxf.flags & 1) if hasattr(ve.dxf, 'flags') else False
                                        if closed and len(pts) >= 3:
                                            if pts[0] != pts[-1]:
                                                pts.append(pts[0])
                                            polygon_records.append({**ve_attrs, 'geometry': Polygon(pts)})
                                        else:
                                            line_records.append({**ve_attrs, 'geometry': LineString(pts)})
                                except Exception:
                                    pass

                            elif ve_type == 'CIRCLE':
                                cx, cy = ve_dxf.center.x, ve_dxf.center.y
                                r = ve_dxf.radius
                                n_seg = self._circle_segments(r, quality)
                                circle_pts = []
                                for i in range(n_seg + 1):
                                    a = 2 * math.pi * i / n_seg
                                    circle_pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
                                polygon_records.append({**ve_attrs, 'geometry': Polygon(circle_pts),
                                                        'radius': r})

                            elif ve_type == 'ARC':
                                cx, cy = ve_dxf.center.x, ve_dxf.center.y
                                radius = ve_dxf.radius
                                sa = math.radians(ve_dxf.start_angle)
                                ea = math.radians(ve_dxf.end_angle)
                                if ea <= sa:
                                    ea += 2 * math.pi
                                arc_span = ea - sa
                                n_seg = max(16, int(64 * arc_span / (2 * math.pi)))
                                pts = []
                                for i in range(n_seg + 1):
                                    a = sa + arc_span * i / n_seg
                                    pts.append((cx + radius * math.cos(a), cy + radius * math.sin(a)))
                                if len(pts) >= 2:
                                    line_records.append({**ve_attrs, 'geometry': LineString(pts),
                                                         'radius': radius})

                            elif ve_type == 'HATCH':
                                try:
                                    for path in ve.paths:
                                        if hasattr(path, 'vertices'):
                                            pts = [(v[0], v[1]) for v in path.vertices]
                                            if len(pts) >= 3:
                                                if pts[0] != pts[-1]:
                                                    pts.append(pts[0])
                                                polygon_records.append({**ve_attrs, 'geometry': Polygon(pts)})
                                except Exception:
                                    pass

                            elif ve_type == 'INSERT':
                                # ネストINSERTを再帰展開（最大2段）
                                try:
                                    for ve2 in ve.virtual_entities():
                                        ve2_type = ve2.dxftype()
                                        ve2_dxf = ve2.dxf
                                        ve2_layer = ve2_dxf.layer if hasattr(ve2_dxf, 'layer') else ve_layer
                                        ve2_color_hex = _resolve_entity_color(ve2, layer_table)
                                        ve2_lt = ve2_dxf.linetype if hasattr(ve2_dxf, 'linetype') else ''
                                        ve2_lw = ve2_dxf.lineweight if hasattr(ve2_dxf, 'lineweight') else -3
                                        if not ve2_lt or ve2_lt in ('ByLayer', 'BYLAYER'):
                                            ve2_lt = layer_table.get(ve2_layer, {}).get('linetype', '実線')
                                        if ve2_lw < 0:
                                            ve2_lw = layer_table.get(ve2_layer, {}).get('lineweight', 25)
                                            if ve2_lw < 0:
                                                ve2_lw = 25
                                        ve2_lw_mm = _dxf_lineweight_mm(ve2_lw)
                                        ve2_dash = _dxf_linetype_to_qgis_dash(ve2_lt, doc)
                                        if scale_factor > 0 and ve2_dash:
                                            ve2_dash = ';'.join(
                                                str(round(float(v) * scale_factor, 4))
                                                for v in ve2_dash.split(';')
                                            )
                                        ve2_pen = 'solid'
                                        if '点線' in ve2_lt:
                                            ve2_pen = 'dot'
                                        elif '二点' in ve2_lt or '三点' in ve2_lt:
                                            ve2_pen = 'dash dot dot'
                                        elif '一点' in ve2_lt or '鎖線' in ve2_lt:
                                            ve2_pen = 'dash dot'
                                        elif '破線' in ve2_lt or '跳び' in ve2_lt:
                                            ve2_pen = 'dash'
                                        elif ve2_dash:
                                            ve2_pen = 'dash'
                                        ve2_attrs = {
                                            'layer': ve2_layer,
                                            'color_idx': ve2_dxf.color if hasattr(ve2_dxf, 'color') else 7,
                                            'color_hex': ve2_color_hex,
                                            'linetype': ve2_lt,
                                            'lineweight': ve2_lw_mm,
                                            'pen_style': ve2_pen,
                                            'dash_pattern': ve2_dash,
                                            'entity_type': ve2_type,
                                        }
                                        if ve2_type == 'LINE':
                                            p1, p2 = ve2_dxf.start, ve2_dxf.end
                                            line_records.append({**ve2_attrs, 'geometry': LineString([(p1.x, p1.y), (p2.x, p2.y)])})
                                        elif ve2_type in ('LWPOLYLINE', 'POLYLINE'):
                                            if ve2_type == 'LWPOLYLINE':
                                                pts = [(p[0], p[1]) for p in ve2.get_points(format='xy')]
                                            else:
                                                pts = [(v.dxf.location.x, v.dxf.location.y) for v in ve2.vertices]
                                            if len(pts) >= 2:
                                                if hasattr(ve2, 'closed'):
                                                    cl = ve2.closed
                                                elif hasattr(ve2, 'is_closed'):
                                                    cl = ve2.is_closed
                                                else:
                                                    cl = bool(ve2.dxf.flags & 1) if hasattr(ve2.dxf, 'flags') else False
                                                if cl and len(pts) >= 3:
                                                    if pts[0] != pts[-1]:
                                                        pts.append(pts[0])
                                                    polygon_records.append({**ve2_attrs, 'geometry': Polygon(pts)})
                                                else:
                                                    line_records.append({**ve2_attrs, 'geometry': LineString(pts)})
                                        elif ve2_type == 'CIRCLE':
                                            cx2, cy2 = ve2_dxf.center.x, ve2_dxf.center.y
                                            r2 = ve2_dxf.radius
                                            ns = self._circle_segments(r2, quality)
                                            cpts = [(cx2 + r2 * math.cos(2 * math.pi * i / ns),
                                                     cy2 + r2 * math.sin(2 * math.pi * i / ns)) for i in range(ns + 1)]
                                            polygon_records.append({**ve2_attrs, 'geometry': Polygon(cpts), 'radius': r2})
                                        elif ve2_type == 'ARC':
                                            cx2, cy2 = ve2_dxf.center.x, ve2_dxf.center.y
                                            r2 = ve2_dxf.radius
                                            sa2 = math.radians(ve2_dxf.start_angle)
                                            ea2 = math.radians(ve2_dxf.end_angle)
                                            if ea2 <= sa2:
                                                ea2 += 2 * math.pi
                                            span2 = ea2 - sa2
                                            ns2 = max(16, int(64 * span2 / (2 * math.pi)))
                                            pts = [(cx2 + r2 * math.cos(sa2 + span2 * i / ns2),
                                                    cy2 + r2 * math.sin(sa2 + span2 * i / ns2)) for i in range(ns2 + 1)]
                                            if len(pts) >= 2:
                                                line_records.append({**ve2_attrs, 'geometry': LineString(pts), 'radius': r2})
                                        elif ve2_type == 'TEXT':
                                            bt = self._clean_text(ve2_dxf.text) if hasattr(ve2_dxf, 'text') else ''
                                            if bt and bt.strip():
                                                bx, by, b_anchor, b_hali, b_vali = self._get_text_anchor_point(ve2_dxf)
                                                bh = ve2_dxf.height if hasattr(ve2_dxf, 'height') else 1.0
                                                br = ve2_dxf.rotation if hasattr(ve2_dxf, 'rotation') else 0.0
                                                ve_texts.append((bx, by, bt, bh, br, b_anchor, b_hali, b_vali, ve2_color_hex))
                                        elif ve2_type == 'HATCH':
                                            try:
                                                for path in ve2.paths:
                                                    if hasattr(path, 'vertices'):
                                                        pts = [(v[0], v[1]) for v in path.vertices]
                                                        if len(pts) >= 3:
                                                            if pts[0] != pts[-1]:
                                                                pts.append(pts[0])
                                                            polygon_records.append({**ve2_attrs, 'geometry': Polygon(pts)})
                                            except Exception:
                                                pass
                                except Exception:
                                    pass

                        # ブロック内TEXTをクラスタリング→グループ毎に結合
                        if ve_texts:
                            # --- Step 1: 空間クラスタリング（近接文字をグループ化）---
                            # ve_texts: (x, y, text, height, rotation, anchor, hali, vali, color_hex)
                            avg_h = sum(t[3] for t in ve_texts) / len(ve_texts)
                            threshold = avg_h * 4.0  # 文字高の4倍以内を同一グループ

                            groups = []  # list of lists
                            assigned = [False] * len(ve_texts)
                            for i in range(len(ve_texts)):
                                if assigned[i]:
                                    continue
                                group = [i]
                                assigned[i] = True
                                queue = [i]
                                while queue:
                                    ci = queue.pop(0)
                                    cx, cy = ve_texts[ci][0], ve_texts[ci][1]
                                    for j in range(len(ve_texts)):
                                        if assigned[j]:
                                            continue
                                        dx = ve_texts[j][0] - cx
                                        dy = ve_texts[j][1] - cy
                                        dist = math.sqrt(dx*dx + dy*dy)
                                        if dist <= threshold:
                                            assigned[j] = True
                                            group.append(j)
                                            queue.append(j)
                                groups.append(group)

                            # --- Step 2: 各グループを適切にソート・結合 ---
                            font = _map_font_family(text_styles.get(ve_style, ''))
                            for grp_indices in groups:
                                grp = [ve_texts[i] for i in grp_indices]
                                if len(grp) == 1:
                                    # 1文字グループ: そのまま
                                    t = grp[0]
                                    combined_text = t[2].strip()
                                    if not combined_text:
                                        continue
                                    geom = Point(t[0], t[1])
                                    scaled_h = t[3] * scale_factor if scale_factor > 0 else t[3]
                                    ch = t[8] if len(t) > 8 else _aci_to_hex(color)
                                    ogr_style = _build_ogr_label_style(
                                        combined_text, font, scaled_h, t[4], ch, t[5])
                                    rec = {**base_attrs, 'geometry': geom,
                                           'entity_type': 'BLOCK_TEXT',
                                           'text': combined_text,
                                           'text_height': scaled_h,
                                           'rotation': t[4],
                                           'font': font,
                                           'style': ve_style,
                                           'attrib_tag': block_name,
                                           'color_hex': ch,
                                           'char_spacing': 0.0,
                                           'char_width_ratio': 0.0,
                                           'hali': t[6],
                                           'vali': t[7],
                                           'OGR_STYLE': ogr_style}
                                    text_records.append(rec)
                                    continue

                                # 配置パターン検出
                                g_h = grp[0][3]  # 代表文字高
                                g_rot = grp[0][4]  # 代表回転角
                                xs = [t[0] for t in grp]
                                ys = [t[1] for t in grp]
                                x_range = max(xs) - min(xs)
                                y_range = max(ys) - min(ys)

                                is_vertical = False
                                joiner = ''  # 結合文字（縦書きなら改行）

                                # 回転角に基づくテキスト方向ベクトル
                                rad = math.radians(g_rot)
                                dir_x = math.cos(rad)
                                dir_y = math.sin(rad)
                                # 垂直方向（上方向）
                                perp_x = -dir_y
                                perp_y = dir_x

                                # 各文字をテキスト方向/垂直方向にプロジェクション
                                projs = []
                                perps = []
                                for t in grp:
                                    dx = t[0] - grp[0][0]
                                    dy = t[1] - grp[0][1]
                                    projs.append(dx * dir_x + dy * dir_y)
                                    perps.append(dx * perp_x + dy * perp_y)

                                proj_range = max(projs) - min(projs) if projs else 0
                                perp_range = max(perps) - min(perps) if perps else 0

                                if proj_range < 0.5 * g_h and perp_range > g_h:
                                    # 縦書き（テキスト方向に広がりなし、垂直方向に広がり）
                                    is_vertical = True
                                    joiner = '\n'
                                    # 垂直方向の降順（上→下）でソート
                                    indexed = list(zip(perps, range(len(grp))))
                                    indexed.sort(key=lambda x: -x[0])
                                    grp = [grp[idx] for _, idx in indexed]
                                elif proj_range >= perp_range:
                                    # テキスト方向に沿って並ぶ（通常の横書き・斜め）
                                    indexed = list(zip(projs, range(len(grp))))
                                    indexed.sort(key=lambda x: x[0])
                                    grp = [grp[idx] for _, idx in indexed]
                                else:
                                    # 垂直方向に広がり（回転つき縦書き）
                                    is_vertical = True
                                    joiner = '\n'
                                    indexed = list(zip(perps, range(len(grp))))
                                    indexed.sort(key=lambda x: -x[0])
                                    grp = [grp[idx] for _, idx in indexed]

                                # === DXF配置を倣った結合テキスト構築 ===
                                if is_vertical:
                                    combined_text = '\n'.join(
                                        t[2] for t in grp).strip()
                                    # 先頭（最上）文字の上端を挿入点とする
                                    orig_vali = (grp[0][7]
                                                 if len(grp[0]) > 7
                                                 else ve_vali)
                                    if orig_vali in ('Bottom', 'Base'):
                                        ins_x = (grp[0][0]
                                                 + g_h * perp_x)
                                        ins_y = (grp[0][1]
                                                 + g_h * perp_y)
                                    elif orig_vali in ('Half', 'Middle'):
                                        ins_x = (grp[0][0]
                                                 + (g_h / 2.0) * perp_x)
                                        ins_y = (grp[0][1]
                                                 + (g_h / 2.0) * perp_y)
                                    else:
                                        ins_x = grp[0][0]
                                        ins_y = grp[0][1]
                                    a_hali = 'Center'
                                    a_vali = 'Top'
                                else:
                                    # --- 各文字のhalignを考慮した左端位置算出 ---
                                    left_projs = []
                                    for t in grp:
                                        dx_t = t[0] - grp[0][0]
                                        dy_t = t[1] - grp[0][1]
                                        proj = dx_t * dir_x + dy_t * dir_y
                                        ch_c = t[2]
                                        ch_hw = (ord(ch_c) < 0x2E80
                                                 or 0xFF61 <= ord(ch_c) <= 0xFF9F)
                                        ch_w = t[3] * (0.5 if ch_hw else 1.0)
                                        t_hali = t[6] if len(t) > 6 else 'Left'
                                        if t_hali == 'Center':
                                            proj -= ch_w / 2.0
                                        elif t_hali == 'Right':
                                            proj -= ch_w
                                        left_projs.append(proj)

                                    # --- ギャップにスペースを挿入してテキスト構築 ---
                                    parts = []
                                    for idx, t in enumerate(grp):
                                        parts.append(t[2])
                                        if idx < len(grp) - 1:
                                            ch_c = t[2]
                                            ch_hw = (ord(ch_c) < 0x2E80
                                                     or 0xFF61 <= ord(ch_c) <= 0xFF9F)
                                            ch_w = t[3] * (0.5 if ch_hw else 1.0)
                                            left_dist = (left_projs[idx + 1]
                                                         - left_projs[idx])
                                            gap = left_dist - ch_w
                                            if gap > ch_w * 0.3:
                                                space_w = t[3] * 0.5
                                                n_sp = max(1, round(gap / space_w))
                                                parts.append(' ' * n_sp)
                                    combined_text = ''.join(parts).strip()

                                    # --- 先頭文字の左端を挿入点とする ---
                                    first_t = grp[0]
                                    f_ch = first_t[2]
                                    f_hw = (ord(f_ch) < 0x2E80
                                            or 0xFF61 <= ord(f_ch) <= 0xFF9F)
                                    f_cw = first_t[3] * (0.5 if f_hw else 1.0)
                                    f_hali = (first_t[6] if len(first_t) > 6
                                              else 'Left')
                                    if f_hali == 'Center':
                                        h_off = -f_cw / 2.0
                                    elif f_hali == 'Right':
                                        h_off = -f_cw
                                    else:
                                        h_off = 0.0
                                    ins_x = first_t[0] + h_off * dir_x
                                    ins_y = first_t[1] + h_off * dir_y
                                    a_hali = 'Left'
                                    a_vali = (first_t[7] if len(first_t) > 7
                                              else ve_vali)

                                if not combined_text:
                                    continue

                                # char_spacing（参考値）
                                cs = 0.0
                                if is_vertical and len(grp) > 1:
                                    # 縦書き: 文字高をchar_spacingに設定
                                    # → QMLでfont_size = text_height
                                    #   (/ 0.88 補正を回避)
                                    cs = (g_h * scale_factor
                                          if scale_factor > 0 else g_h)
                                elif not is_vertical and len(grp) > 1:
                                    sorted_projs = []
                                    for t in grp:
                                        dx_t = t[0] - grp[0][0]
                                        dy_t = t[1] - grp[0][1]
                                        sorted_projs.append(
                                            dx_t * dir_x + dy_t * dir_y)
                                    spacings = [
                                        sorted_projs[i+1] - sorted_projs[i]
                                        for i in range(len(sorted_projs)-1)]
                                    avg_sp = (sum(spacings) / len(spacings)
                                              if spacings else 0)
                                    cs = (avg_sp * scale_factor
                                          if scale_factor > 0 else avg_sp)
                                all_chars = ''.join(t[2] for t in grp)
                                hw = sum(1 for ch_c in all_chars
                                         if ord(ch_c) < 0x2E80
                                         or 0xFF61 <= ord(ch_c) <= 0xFF9F)
                                cwr = 0.5 if hw > len(all_chars) - hw else 1.0

                                first = grp[0]
                                geom = Point(ins_x, ins_y)
                                scaled_h = (first[3] * scale_factor
                                            if scale_factor > 0 else first[3])
                                ch = (first[8] if len(first) > 8
                                      else _aci_to_hex(color))
                                a_anchor = (first[5] if len(first) > 5
                                            else ve_anchor)
                                ogr_style = _build_ogr_label_style(
                                    combined_text, font, scaled_h,
                                    first[4], ch, a_anchor)
                                rec = {**base_attrs, 'geometry': geom,
                                       'entity_type': 'BLOCK_TEXT',
                                       'text': combined_text,
                                       'text_height': scaled_h,
                                       'rotation': first[4],
                                       'font': font,
                                       'style': ve_style,
                                       'attrib_tag': block_name,
                                       'color_hex': ch,
                                       'char_spacing': cs,
                                       'char_width_ratio': cwr if cs > 0 else 0.0,
                                       'hali': a_hali,
                                       'vali': a_vali,
                                       'OGR_STYLE': ogr_style}
                                text_records.append(rec)
                    except Exception:
                        pass

                elif etype == 'SPLINE':
                    pts = [(p.x, p.y) for p in entity.control_points]
                    if len(pts) >= 2:
                        # スプラインのフィットポイントを使って近似
                        try:
                            flattened = list(entity.flattening(0.5))
                            pts = [(p.x, p.y) for p in flattened]
                        except:
                            pass
                        if len(pts) >= 2:
                            geom = LineString(pts)
                            line_records.append({**base_attrs, 'geometry': geom})

                elif etype == 'HATCH':
                    try:
                        paths = entity.paths
                        for path in paths:
                            if hasattr(path, 'vertices'):
                                pts = [(v[0], v[1]) for v in path.vertices]
                                if len(pts) >= 3:
                                    if pts[0] != pts[-1]:
                                        pts.append(pts[0])
                                    geom = Polygon(pts)
                                    polygon_records.append({**base_attrs, 'geometry': geom})
                    except:
                        pass


            except Exception as e:
                self.warnings.append(f"エンティティ変換エラー ({entity.dxftype()}): {str(e)[:60]}")

        result = {}
        if point_records:
            result['points'] = gpd.GeoDataFrame(point_records, geometry='geometry')
        if line_records:
            result['lines'] = gpd.GeoDataFrame(line_records, geometry='geometry')
        if polygon_records:
            result['polygons'] = gpd.GeoDataFrame(polygon_records, geometry='geometry')
        if text_records:
            result['text'] = gpd.GeoDataFrame(text_records, geometry='geometry')

        # ジオメトリ座標変換: 方眼線GCPからアフィン変換 or スケーリング
        affine_applied = False
        if auto_georef and 'text' in result and scale_factor > 0:
            affine_matrix = self._compute_affine_from_grid(result['text'])
            if affine_matrix is not None:
                a, b, tx, c, d, ty = affine_matrix
                # a,b,tx: Real_X = a*dxf_x + b*dxf_y + tx (JPC_X = 北距 northing)
                # c,d,ty: Real_Y = c*dxf_x + d*dxf_y + ty (JPC_Y = 東距 easting)
                # GIS規約: geometry.x = easting, geometry.y = northing
                # → X/Y軸を入れ替えて適用
                from shapely.affinity import affine_transform as shapely_affine
                for key, gdf in result.items():
                    result[key] = gdf.copy()
                    # shapely affine_transform: [a, b, d, e, xoff, yoff]
                    # new_x = easting  = c*dxf_x + d*dxf_y + ty
                    # new_y = northing = a*dxf_x + b*dxf_y + tx
                    result[key]['geometry'] = gdf['geometry'].apply(
                        lambda g: shapely_affine(g, [c, d, a, b, ty, tx])
                    )
                    if 'radius' in result[key].columns:
                        result[key]['radius'] = result[key]['radius'] * scale_factor

                # テキスト回転角をアフィン変換に合わせて補正
                # GIS座標系でのDXF X軸方向: atan2(a, c) (CCW from east)
                # DXF回転θのテキスト方向ベクトル(cosθ,sinθ)を
                # アフィン変換すると GIS角 = offset + dxf_rot
                if 'text' in result:
                    rotation_offset = math.degrees(math.atan2(a, c))
                    text_gdf = result['text']
                    if 'rotation' in text_gdf.columns:
                        result['text'] = text_gdf.copy()
                        result['text']['rotation'] = rotation_offset + text_gdf['rotation']
                        # [-180, 180) に正規化
                        result['text']['rotation'] = (
                            (result['text']['rotation'] + 180) % 360
                        ) - 180
                    self.warnings.append(
                        f"テキスト回転補正: offset={rotation_offset:.2f}° "
                        f"(DXF X軸→GIS東方向)")

                affine_applied = True
                self.warnings.append(
                    f"方眼線GCPからアフィン変換適用済み "
                    f"(scale={math.sqrt(a*a+b*b):.6f}, "
                    f"rotation={math.degrees(math.atan2(a, b)):.2f}°, "
                    f"JPC軸→GIS軸変換: X=easting(Y_jpc), Y=northing(X_jpc))")

        if not affine_applied and scale_factor > 0:
            from shapely.affinity import scale as shapely_scale
            for key, gdf in result.items():
                result[key] = gdf.copy()
                result[key]['geometry'] = gdf['geometry'].apply(
                    lambda g: shapely_scale(g, xfact=scale_factor, yfact=scale_factor, origin=(0, 0))
                )
                if 'radius' in result[key].columns:
                    result[key]['radius'] = result[key]['radius'] * scale_factor

        # 軽量モード: ジオメトリ簡略化＋不要属性削減
        if quality < 50:
            # Douglas-Peucker簡略化（tolerance = 品質に反比例）
            tolerance = 0.05 if quality <= 30 else 0.02  # メートル単位
            for key in ('lines', 'polygons'):
                if key in result:
                    gdf = result[key]
                    result[key] = gdf.copy()
                    result[key]['geometry'] = gdf['geometry'].simplify(
                        tolerance, preserve_topology=True)
            # 冗長属性カラムを削除
            drop_cols = ['dash_pattern', 'color_idx', 'linetype']
            for key, gdf in result.items():
                existing = [c for c in drop_cols if c in gdf.columns]
                if existing:
                    result[key] = gdf.drop(columns=existing)

        return result


# ============================================================
# GeoPackage 書き出し
# ============================================================
def save_to_geopackage(
    gdfs: dict[str, gpd.GeoDataFrame],
    output_path: str,
    source_crs_epsg: int,
    target_crs_epsg: Optional[int] = None,
    split_by_layer: bool = False,
) -> tuple[bool, list[str]]:
    """GeoDataFrame群をGeoPackageに保存する

    Args:
        split_by_layer: Trueの場合、DXFレイヤー名ごとに分割して
                        GeoPackageレイヤーを作成する
    """
    messages = []

    if not gdfs:
        return False, ["変換可能なデータがありません"]

    try:
        source_crs = CRS.from_epsg(source_crs_epsg)
    except Exception as e:
        return False, [f"ソースCRS設定エラー (EPSG:{source_crs_epsg}): {e}"]

    target_crs = None
    if target_crs_epsg and target_crs_epsg != source_crs_epsg:
        try:
            target_crs = CRS.from_epsg(target_crs_epsg)
        except Exception as e:
            messages.append(f"ターゲットCRS設定エラー: {e}、ソースCRSのまま保存します")
            target_crs = None

    # DXFレイヤー分割モード: layerカラムで分割
    if split_by_layer:
        geom_suffix = {'lines': '_line', 'polygons': '_polygon', 'text': '_text'}
        split_gdfs = {}
        for geom_type, gdf in gdfs.items():
            if gdf.empty or 'layer' not in gdf.columns:
                split_gdfs[geom_type] = gdf
                continue
            suffix = geom_suffix.get(geom_type, f'_{geom_type}')
            for dxf_layer, sub_gdf in gdf.groupby('layer'):
                gpkg_name = f"{dxf_layer}{suffix}"
                split_gdfs[gpkg_name] = sub_gdf
        gdfs = split_gdfs

    first_layer = True
    total_features = 0

    for layer_name, gdf in gdfs.items():
        if gdf.empty:
            continue

        gdf = gdf.copy()
        gdf.set_crs(source_crs, inplace=True)

        if target_crs:
            gdf = gdf.to_crs(target_crs)

        mode = 'w' if first_layer else 'a'
        try:
            gdf.to_file(
                output_path,
                layer=layer_name,
                driver='GPKG',
                mode=mode,
            )
            total_features += len(gdf)
            messages.append(f"  レイヤ '{layer_name}': {len(gdf)} フィーチャ")
            first_layer = False
        except Exception as e:
            messages.append(f"  レイヤ '{layer_name}' 書き込みエラー: {e}")

    if total_features > 0:
        # 全レイヤーにQGISスタイルを埋め込む
        try:
            _embed_qgis_styles(output_path, gdfs)
            messages.append("  QGISスタイル埋め込み済み")
        except Exception as e:
            messages.append(f"  スタイル埋め込みスキップ: {e}")

        messages.insert(0, f"合計 {total_features} フィーチャを保存しました")
        return True, messages
    else:
        return False, ["書き込み可能なフィーチャがありません"]


def _embed_qgis_styles(gpkg_path: str, gdfs: dict):
    """GeoPackageに全レイヤーのQGISデフォルトスタイルを埋め込む

    - lines: color_hex/lineweight/dash_patternでデータ駆動スタイル
    - polygons: 同上 + 塗りなし
    - points: color_hexでマーカー色
    - text: テキストラベル自動表示（Size/Rotation/Hali/Vali）
    """
    import sqlite3

    # ======== ラインレイヤー用QML ========
    # color_hex, lineweight, pen_style でデータ駆動描画
    line_qml = '''<!DOCTYPE qgis PUBLIC 'http://mrcc.com/qgis.dtd' 'SYSTEM'>
<qgis styleCategories="Symbology" version="3.34">
  <renderer-v2 type="singleSymbol">
    <symbols>
      <symbol type="line" name="0" alpha="1" force_rhr="0" clip_to_extent="1">
        <data_defined_properties>
          <Option type="Map">
            <Option type="QString" name="name" value=""/>
            <Option type="Map" name="properties"/>
            <Option type="QString" name="type" value="collection"/>
          </Option>
        </data_defined_properties>
        <layer class="SimpleLine" pass="0" locked="0" enabled="1">
          <Option type="Map">
            <Option type="QString" name="line_color" value="0,0,0,255"/>
            <Option type="QString" name="line_style" value="solid"/>
            <Option type="QString" name="line_width" value="0.25"/>
            <Option type="QString" name="line_width_unit" value="MM"/>
            <Option type="QString" name="capstyle" value="round"/>
            <Option type="QString" name="joinstyle" value="round"/>
          </Option>
          <data_defined_properties>
            <Option type="Map">
              <Option type="QString" name="name" value=""/>
              <Option type="Map" name="properties">
                <Option type="Map" name="outlineColor">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="field" value="color_hex"/>
                  <Option type="int" name="type" value="2"/>
                </Option>
                <Option type="Map" name="outlineWidth">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="field" value="lineweight"/>
                  <Option type="int" name="type" value="2"/>
                </Option>
                <Option type="Map" name="outlineStyle">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="field" value="pen_style"/>
                  <Option type="int" name="type" value="2"/>
                </Option>
              </Option>
              <Option type="QString" name="type" value="collection"/>
            </Option>
          </data_defined_properties>
        </layer>
      </symbol>
    </symbols>
  </renderer-v2>
</qgis>'''

    # ======== ポリゴンレイヤー用QML ========
    polygon_qml = '''<!DOCTYPE qgis PUBLIC 'http://mrcc.com/qgis.dtd' 'SYSTEM'>
<qgis styleCategories="Symbology" version="3.34">
  <renderer-v2 type="singleSymbol">
    <symbols>
      <symbol type="fill" name="0" alpha="1" force_rhr="0" clip_to_extent="1">
        <data_defined_properties>
          <Option type="Map">
            <Option type="QString" name="name" value=""/>
            <Option type="Map" name="properties"/>
            <Option type="QString" name="type" value="collection"/>
          </Option>
        </data_defined_properties>
        <layer class="SimpleFill" pass="0" locked="0" enabled="1">
          <Option type="Map">
            <Option type="QString" name="color" value="0,0,0,0"/>
            <Option type="QString" name="style" value="no"/>
            <Option type="QString" name="outline_color" value="0,0,0,255"/>
            <Option type="QString" name="outline_style" value="solid"/>
            <Option type="QString" name="outline_width" value="0.25"/>
            <Option type="QString" name="outline_width_unit" value="MM"/>
          </Option>
          <data_defined_properties>
            <Option type="Map">
              <Option type="QString" name="name" value=""/>
              <Option type="Map" name="properties">
                <Option type="Map" name="outlineColor">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="field" value="color_hex"/>
                  <Option type="int" name="type" value="2"/>
                </Option>
                <Option type="Map" name="outlineWidth">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="field" value="lineweight"/>
                  <Option type="int" name="type" value="2"/>
                </Option>
                <Option type="Map" name="outlineStyle">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="field" value="pen_style"/>
                  <Option type="int" name="type" value="2"/>
                </Option>
              </Option>
              <Option type="QString" name="type" value="collection"/>
            </Option>
          </data_defined_properties>
        </layer>
      </symbol>
    </symbols>
  </renderer-v2>
</qgis>'''

    # ======== ポイントレイヤー用QML ========
    # CIRCLEはradiusでMapUnit径指定、POINTは固定2mm表示
    point_qml = '''<!DOCTYPE qgis PUBLIC 'http://mrcc.com/qgis.dtd' 'SYSTEM'>
<qgis styleCategories="Symbology" version="3.34">
  <renderer-v2 type="singleSymbol">
    <symbols>
      <symbol type="marker" name="0" alpha="1" force_rhr="0" clip_to_extent="1">
        <data_defined_properties>
          <Option type="Map">
            <Option type="QString" name="name" value=""/>
            <Option type="Map" name="properties"/>
            <Option type="QString" name="type" value="collection"/>
          </Option>
        </data_defined_properties>
        <layer class="SimpleMarker" pass="0" locked="0" enabled="1">
          <Option type="Map">
            <Option type="QString" name="name" value="circle"/>
            <Option type="QString" name="color" value="0,0,0,255"/>
            <Option type="QString" name="outline_color" value="0,0,0,255"/>
            <Option type="QString" name="size" value="2"/>
            <Option type="QString" name="size_unit" value="MM"/>
            <Option type="QString" name="outline_style" value="solid"/>
            <Option type="QString" name="outline_width" value="0.15"/>
            <Option type="QString" name="outline_width_unit" value="MM"/>
          </Option>
          <data_defined_properties>
            <Option type="Map">
              <Option type="QString" name="name" value=""/>
              <Option type="Map" name="properties">
                <Option type="Map" name="fillColor">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="field" value="color_hex"/>
                  <Option type="int" name="type" value="2"/>
                </Option>
                <Option type="Map" name="outlineColor">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="field" value="color_hex"/>
                  <Option type="int" name="type" value="2"/>
                </Option>
                <Option type="Map" name="size">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="expression"
                    value="if(&quot;radius&quot; IS NOT NULL AND &quot;radius&quot; > 0, &quot;radius&quot; * 2, 2)"/>
                  <Option type="int" name="type" value="3"/>
                </Option>
                <Option type="Map" name="sizeUnit">
                  <Option type="bool" name="active" value="true"/>
                  <Option type="QString" name="expression"
                    value="if(&quot;radius&quot; IS NOT NULL AND &quot;radius&quot; > 0, 'MapUnit', 'MM')"/>
                  <Option type="int" name="type" value="3"/>
                </Option>
              </Option>
              <Option type="QString" name="type" value="collection"/>
            </Option>
          </data_defined_properties>
        </layer>
      </symbol>
    </symbols>
  </renderer-v2>
</qgis>'''

    # ======== テキストレイヤー用QML ========
    text_qml = '''<!DOCTYPE qgis PUBLIC 'http://mrcc.com/qgis.dtd' 'SYSTEM'>
<qgis styleCategories="Symbology|Labeling" version="3.34">
  <renderer-v2 type="nullSymbol"/>
  <labeling type="simple">
    <settings calloutType="simple">
      <text-style fieldName="text" fontSize="10" fontSizeUnit="MapUnit"
                  textColor="0,0,0,255"
                  fontFamily="Noto Sans CJK JP" fontWeight="50"
                  namedStyle="Regular" allowHtml="0" multilineHeight="1">
        <text-buffer bufferDraw="0"/>
        <text-mask maskEnabled="0"/>
        <background shapeDraw="0"/>
        <shadow shadowDraw="0"/>
        <dd_properties>
          <Option type="Map">
            <Option type="QString" name="name" value=""/>
            <Option type="Map" name="properties">
              <Option type="Map" name="Color">
                <Option type="bool" name="active" value="true"/>
                <Option type="QString" name="field" value="color_hex"/>
                <Option type="int" name="type" value="2"/>
              </Option>
              <Option type="Map" name="Family">
                <Option type="bool" name="active" value="true"/>
                <Option type="QString" name="field" value="font"/>
                <Option type="int" name="type" value="2"/>
              </Option>
            </Option>
            <Option type="QString" name="type" value="collection"/>
          </Option>
        </dd_properties>
      </text-style>
      <text-format multilineAlign="0" wrapChar="" autoWrapLength="0"
                   addDirectionSymbol="0" reverseDirectionSymbol="0"
                   useMaxLineLengthForAutoWrap="1" placeDirectionSymbol="0"/>
      <placement placement="1" priority="10" quadOffset="4"
                 xOffset="0" yOffset="0"
                 rotationAngle="0" rotationUnit="AngleDegrees"
                 preserveRotation="1"
                 dist="0" distUnits="MM"
                 geometryGeneratorEnabled="0" layerType="PointGeometry"/>
      <rendering displayAll="1" scaleVisibility="0"
                 fontMinPixelSize="3" fontLimitPixelSize="0" fontMaxPixelSize="10000"
                 labelPerPart="0" mergeLines="0"
                 obstacle="0" obstacleFactor="1" obstacleType="1"
                 limitNumLabels="0" maxNumLabels="2000"
                 drawUnplacedLabels="1" upsidedownLabels="0" zIndex="0"/>
      <dd_properties>
        <Option type="Map">
          <Option type="QString" name="name" value=""/>
          <Option type="Map" name="properties">
            <Option type="Map" name="Size">
              <Option type="bool" name="active" value="true"/>
              <Option type="QString" name="expression" value="CASE WHEN &quot;char_spacing&quot; &gt; 0 THEN &quot;text_height&quot; ELSE &quot;text_height&quot; / 0.88 END"/>
              <Option type="int" name="type" value="3"/>
            </Option>
            <Option type="Map" name="LabelRotation">
              <Option type="bool" name="active" value="true"/>
              <Option type="QString" name="expression" value="-&quot;rotation&quot;"/>
              <Option type="int" name="type" value="3"/>
            </Option>
            <Option type="Map" name="Hali">
              <Option type="bool" name="active" value="true"/>
              <Option type="QString" name="field" value="hali"/>
              <Option type="int" name="type" value="2"/>
            </Option>
            <Option type="Map" name="Vali">
              <Option type="bool" name="active" value="true"/>
              <Option type="QString" name="field" value="vali"/>
              <Option type="int" name="type" value="2"/>
            </Option>
            <Option type="Map" name="OffsetQuad">
              <Option type="bool" name="active" value="true"/>
              <Option type="QString" name="expression" value="CASE WHEN &quot;hali&quot; = 'Left' AND &quot;vali&quot; IN ('Bottom','Base') THEN 2 WHEN &quot;hali&quot; = 'Left' AND &quot;vali&quot; = 'Half' THEN 5 WHEN &quot;hali&quot; = 'Left' AND &quot;vali&quot; = 'Top' THEN 8 WHEN &quot;hali&quot; = 'Center' AND &quot;vali&quot; IN ('Bottom','Base') THEN 1 WHEN &quot;hali&quot; = 'Center' AND &quot;vali&quot; = 'Half' THEN 4 WHEN &quot;hali&quot; = 'Center' AND &quot;vali&quot; = 'Top' THEN 7 WHEN &quot;hali&quot; = 'Right' AND &quot;vali&quot; IN ('Bottom','Base') THEN 0 WHEN &quot;hali&quot; = 'Right' AND &quot;vali&quot; = 'Half' THEN 3 WHEN &quot;hali&quot; = 'Right' AND &quot;vali&quot; = 'Top' THEN 6 ELSE 4 END"/>
              <Option type="int" name="type" value="3"/>
            </Option>
          </Option>
          <Option type="QString" name="type" value="collection"/>
        </Option>
      </dd_properties>
    </settings>
  </labeling>
</qgis>'''

    # ジオメトリ種別 → QMLのマッピング
    style_map = {
        'lines': ('ライン描画（色・線幅・線種）', line_qml),
        'polygons': ('ポリゴン描画（枠線色・線幅）', polygon_qml),
        'points': ('ポイント描画（色・サイズ）', point_qml),
        'text': ('テキストラベル自動表示', text_qml),
    }
    # サフィックスによるマッチ（split_by_layer用）
    suffix_map = {
        '_line': ('ライン描画', line_qml),
        '_polygon': ('ポリゴン描画', polygon_qml),
        '_point': ('ポイント描画', point_qml),
        '_text': ('テキストラベル', text_qml),
    }

    conn = sqlite3.connect(gpkg_path)
    try:
        # 既存テーブルを削除して再作成（重複防止）
        conn.execute('DROP TABLE IF EXISTS layer_styles')
        conn.execute('''CREATE TABLE layer_styles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            f_table_catalog TEXT DEFAULT '',
            f_table_schema TEXT DEFAULT '',
            f_table_name TEXT NOT NULL,
            f_geometry_column TEXT,
            styleName TEXT DEFAULT 'default',
            styleQML TEXT,
            styleSLD TEXT,
            useAsDefault BOOLEAN DEFAULT 1,
            description TEXT,
            owner TEXT DEFAULT '',
            ui TEXT,
            update_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )''')

        for layer_name, gdf in gdfs.items():
            if gdf.empty:
                continue

            # レイヤー名 → QMLを特定
            desc, qml = None, None
            if layer_name in style_map:
                desc, qml = style_map[layer_name]
            else:
                for suffix, (s_desc, s_qml) in suffix_map.items():
                    if layer_name.endswith(suffix):
                        desc, qml = s_desc, s_qml
                        break
            if qml is None:
                continue

            # GeoPackageから実際のジオメトリカラム名を取得
            cursor = conn.execute(
                "SELECT column_name FROM gpkg_geometry_columns WHERE table_name = ?",
                (layer_name,)
            )
            row = cursor.fetchone()
            geom_col = row[0] if row else 'geometry'

            conn.execute(
                '''INSERT INTO layer_styles
                   (f_table_name, f_geometry_column, styleName, styleQML, useAsDefault, description)
                   VALUES (?, ?, 'default', ?, 1, ?)''',
                (layer_name, geom_col, qml, desc)
            )

        conn.commit()
    finally:
        conn.close()


# ============================================================
# GUI - Tkinter Application
# ============================================================
class ConverterApp:
    """SFC/DXF → GeoPackage 変換 GUIアプリ"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("SFC/DXF → GeoPackage Converter")
        self.root.geometry("780x780")
        self.root.resizable(True, True)

        # 変数
        self.input_files: list[str] = []
        self.output_path = tk.StringVar()
        self.source_datum = tk.StringVar(value="JGD2011")
        self.zone_number = tk.StringVar(value="9")
        self.target_crs = tk.StringVar(value="same")  # same / epsg4326 / custom
        self.custom_epsg = tk.StringVar(value="")
        self.merge_layers = tk.BooleanVar(value=False)

        # DXF変換設定
        self.scale_denominator = tk.StringVar(value="0")
        self.quality = tk.StringVar(value="100")
        self.auto_georef = tk.BooleanVar(value=True)
        self.split_by_layer = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self):
        """UIを構築"""
        # メインフレーム
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # --- 入力ファイル ---
        lf_input = ttk.LabelFrame(main, text="① 入力ファイル（SFC / DXF）", padding=8)
        lf_input.pack(fill=tk.X, pady=(0, 8))

        btn_frame = ttk.Frame(lf_input)
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="ファイル追加...", command=self._add_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="フォルダ追加...", command=self._add_folder).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="クリア", command=self._clear_files).pack(side=tk.LEFT)

        self.file_listbox = tk.Listbox(lf_input, height=5, selectmode=tk.EXTENDED)
        self.file_listbox.pack(fill=tk.X, pady=(5, 0))

        # --- 座標系設定 ---
        lf_crs = ttk.LabelFrame(main, text="② 座標系設定", padding=8)
        lf_crs.pack(fill=tk.X, pady=(0, 8))

        # ソース座標系
        src_frame = ttk.Frame(lf_crs)
        src_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(src_frame, text="測地系:").pack(side=tk.LEFT)
        ttk.Radiobutton(src_frame, text="JGD2011", variable=self.source_datum,
                        value="JGD2011").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(src_frame, text="JGD2000", variable=self.source_datum,
                        value="JGD2000").pack(side=tk.LEFT, padx=5)

        zone_frame = ttk.Frame(lf_crs)
        zone_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(zone_frame, text="系番号:").pack(side=tk.LEFT)
        zone_combo = ttk.Combobox(
            zone_frame, textvariable=self.zone_number, width=50,
            values=[f"{k} - {v}" for k, v in ZONE_DESCRIPTIONS.items()],
            state='readonly'
        )
        zone_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        zone_combo.set("9 - IX系 (東京都,福島,栃木,茨城,埼玉,千葉,群馬,神奈川)")
        zone_combo.bind('<<ComboboxSelected>>', self._on_zone_selected)

        # ターゲット座標系
        tgt_frame = ttk.Frame(lf_crs)
        tgt_frame.pack(fill=tk.X)

        ttk.Label(tgt_frame, text="出力CRS:").pack(side=tk.LEFT)
        ttk.Radiobutton(tgt_frame, text="ソースと同じ", variable=self.target_crs,
                        value="same").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(tgt_frame, text="WGS84 (緯度経度)", variable=self.target_crs,
                        value="epsg4326").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(tgt_frame, text="EPSG指定:", variable=self.target_crs,
                        value="custom").pack(side=tk.LEFT, padx=5)
        ttk.Entry(tgt_frame, textvariable=self.custom_epsg, width=8).pack(side=tk.LEFT)

        # --- DXF変換設定 ---
        lf_dxf = ttk.LabelFrame(main, text="③ DXF変換設定", padding=8)
        lf_dxf.pack(fill=tk.X, pady=(0, 8))

        # 縮尺設定
        scale_frame = ttk.Frame(lf_dxf)
        scale_frame.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(scale_frame, text="図面縮尺:").pack(side=tk.LEFT)
        scale_combo = ttk.Combobox(
            scale_frame, textvariable=self.scale_denominator, width=12,
            values=["0 (スケーリング無し)", "200", "250", "300", "500", "1000", "2500", "5000"],
        )
        scale_combo.pack(side=tk.LEFT, padx=5)
        scale_combo.set("0 (スケーリング無し)")
        ttk.Label(scale_frame, text="（例: 1/300図面 → 300）").pack(side=tk.LEFT)

        # 品質設定
        qual_frame = ttk.Frame(lf_dxf)
        qual_frame.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(qual_frame, text="出力品質:").pack(side=tk.LEFT)
        ttk.Radiobutton(qual_frame, text="100% (高精細)",
                        variable=self.quality, value="100").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(qual_frame, text="50% (標準)",
                        variable=self.quality, value="50").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(qual_frame, text="30% (軽量)",
                        variable=self.quality, value="30").pack(side=tk.LEFT, padx=5)

        # オプションチェックボックス
        opt_dxf_frame = ttk.Frame(lf_dxf)
        opt_dxf_frame.pack(fill=tk.X)
        ttk.Checkbutton(opt_dxf_frame, text="方眼線から自動ジオリファレンス",
                        variable=self.auto_georef).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Checkbutton(opt_dxf_frame, text="DXFレイヤ構造を保持",
                        variable=self.split_by_layer).pack(side=tk.LEFT)

        # --- 出力設定 ---
        lf_out = ttk.LabelFrame(main, text="④ 出力設定", padding=8)
        lf_out.pack(fill=tk.X, pady=(0, 8))

        out_frame = ttk.Frame(lf_out)
        out_frame.pack(fill=tk.X)
        ttk.Label(out_frame, text="出力先:").pack(side=tk.LEFT)
        ttk.Entry(out_frame, textvariable=self.output_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(out_frame, text="参照...", command=self._browse_output).pack(side=tk.LEFT)

        opt_frame = ttk.Frame(lf_out)
        opt_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Checkbutton(opt_frame, text="全ファイルを1つのGeoPackageに統合",
                        variable=self.merge_layers).pack(side=tk.LEFT)

        # --- 変換ボタン ---
        btn_convert = ttk.Button(main, text="🔄 変換実行", command=self._run_conversion)
        btn_convert.pack(pady=8)

        # --- ログ表示 ---
        lf_log = ttk.LabelFrame(main, text="ログ", padding=8)
        lf_log.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(lf_log, height=12, wrap=tk.WORD, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(lf_log, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _log(self, msg: str):
        """ログにメッセージ追加"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.root.update_idletasks()

    def _on_zone_selected(self, event):
        """系番号Comboboxの選択イベント"""
        sel = event.widget.get()
        zone = int(sel.split(' - ')[0])
        self.zone_number.set(zone)

    def _analyze_file(self, filepath: str):
        """ファイルの座標範囲を解析してログに表示"""
        ext = Path(filepath).suffix.lower()
        fname = Path(filepath).name

        self._log(f"\n📐 座標解析: {fname}")

        if ext == '.dxf':
            info = analyze_dxf_coordinates(filepath)
        elif ext == '.sfc':
            info = analyze_sfc_coordinates(filepath)
        else:
            self._log("  未対応形式")
            return

        if info.get('error'):
            self._log(f"  エラー: {info['error']}")
            return

        if info['x_min'] is not None:
            self._log(f"  X範囲: {info['x_min']:.1f} ～ {info['x_max']:.1f}")
            self._log(f"  Y範囲: {info['y_min']:.1f} ～ {info['y_max']:.1f}")
            self._log(f"  エンティティ数: {info['entity_count']}")
            if info['layer_names']:
                layers_str = ', '.join(sorted(info['layer_names'])[:10])
                if len(info['layer_names']) > 10:
                    layers_str += f' ... 他{len(info["layer_names"])-10}レイヤ'
                self._log(f"  レイヤ: {layers_str}")
            self._log(f"  🔍 {info['coord_type']}")

            candidates = info.get('zone_candidates', [])
            if candidates:
                self._log("  💡 作業地域を見て、正しい系番号を選んでください")
                if info['suggested_zone']:
                    zone = info['suggested_zone']
                    self.zone_number.set(zone)
                    for widget in self.root.winfo_children():
                        self._update_zone_combo(widget, zone)
        else:
            self._log("  座標データが見つかりません")

    def _update_zone_combo(self, widget, zone):
        """再帰的にComboboxを探して系番号を更新"""
        if isinstance(widget, ttk.Combobox):
            values = widget.cget('values')
            if values and '系' in str(values[0]):
                widget.set(f"{zone} - {ZONE_DESCRIPTIONS[zone]}")
                return
        for child in widget.winfo_children():
            self._update_zone_combo(child, zone)

    def _add_files(self):
        """ファイル選択ダイアログ"""
        files = filedialog.askopenfilenames(
            title="SFC/DXFファイルを選択",
            filetypes=[
                ("CAD files", "*.sfc *.SFC *.dxf *.DXF"),
                ("SFC files", "*.sfc *.SFC"),
                ("DXF files", "*.dxf *.DXF"),
                ("All files", "*.*"),
            ]
        )
        for f in files:
            if f not in self.input_files:
                self.input_files.append(f)
                self.file_listbox.insert(tk.END, Path(f).name)
                self._analyze_file(f)

    def _add_folder(self):
        """フォルダ内のSFC/DXFを一括追加"""
        folder = filedialog.askdirectory(title="フォルダを選択")
        if folder:
            count = 0
            for f in Path(folder).glob('**/*'):
                if f.suffix.lower() in ('.sfc', '.dxf'):
                    fpath = str(f)
                    if fpath not in self.input_files:
                        self.input_files.append(fpath)
                        self.file_listbox.insert(tk.END, f.name)
                        count += 1
            self._log(f"{count} ファイルを追加しました")
            # 最初のファイルだけ座標解析
            if self.input_files:
                self._analyze_file(self.input_files[-1])

    def _clear_files(self):
        """ファイルリストクリア"""
        self.input_files.clear()
        self.file_listbox.delete(0, tk.END)

    def _browse_output(self):
        """出力先参照"""
        if self.merge_layers.get():
            path = filedialog.asksaveasfilename(
                title="出力GeoPackage",
                defaultextension=".gpkg",
                filetypes=[("GeoPackage", "*.gpkg")]
            )
        else:
            path = filedialog.askdirectory(title="出力フォルダを選択")
        if path:
            self.output_path.set(path)

    def _get_source_epsg(self) -> int:
        """ソースEPSGコードを取得"""
        zone_str = self.zone_number.get()
        try:
            zone = int(zone_str.split(' ')[0].split('-')[0].strip())
        except (ValueError, IndexError):
            zone = 9  # デフォルト
        if self.source_datum.get() == "JGD2011":
            return JGD2011_EPSG.get(zone, 6677)
        else:
            return JGD2000_EPSG.get(zone, 2451)

    def _get_target_epsg(self) -> Optional[int]:
        """ターゲットEPSGコードを取得"""
        tgt = self.target_crs.get()
        if tgt == "same":
            return None
        elif tgt == "epsg4326":
            return 4326
        elif tgt == "custom":
            try:
                return int(self.custom_epsg.get())
            except ValueError:
                return None
        return None

    def _run_conversion(self):
        """変換実行"""
        if not self.input_files:
            messagebox.showwarning("注意", "入力ファイルを選択してください")
            return

        output = self.output_path.get()
        if not output:
            messagebox.showwarning("注意", "出力先を指定してください")
            return

        source_epsg = self._get_source_epsg()
        target_epsg = self._get_target_epsg()

        scale_denom = self._get_scale_denominator()
        quality = self._get_quality()

        self._log("=" * 60)
        self._log(f"変換開始...")
        self._log(f"ソースCRS: EPSG:{source_epsg}")
        if target_epsg:
            self._log(f"ターゲットCRS: EPSG:{target_epsg}")
        if scale_denom > 0:
            self._log(f"図面縮尺: 1/{scale_denom}")
        self._log(f"品質: {quality}%  |  ジオリファレンス: "
                  f"{'ON' if self.auto_georef.get() else 'OFF'}  |  "
                  f"レイヤ分割: {'ON' if self.split_by_layer.get() else 'OFF'}")

        merge = self.merge_layers.get()
        all_gdfs: dict[str, gpd.GeoDataFrame] = {}

        for filepath in self.input_files:
            fname = Path(filepath).name
            ext = Path(filepath).suffix.lower()
            self._log(f"\n--- {fname} ---")

            try:
                if ext == '.sfc':
                    gdfs = self._convert_sfc(filepath)
                elif ext == '.dxf':
                    gdfs = self._convert_dxf(filepath)
                else:
                    self._log(f"  未対応形式: {ext}")
                    continue

                if not gdfs:
                    self._log("  変換可能なフィーチャがありません")
                    continue

                if merge:
                    # 統合モード: プレフィックスを付けてマージ
                    stem = Path(filepath).stem
                    for lname, gdf in gdfs.items():
                        key = f"{stem}_{lname}"
                        all_gdfs[key] = gdf
                        self._log(f"  {lname}: {len(gdf)} フィーチャ")
                else:
                    # 個別保存
                    out_path = os.path.join(output, Path(filepath).stem + ".gpkg")
                    success, msgs = save_to_geopackage(
                        gdfs, out_path, source_epsg, target_epsg,
                        split_by_layer=self.split_by_layer.get(),
                    )
                    for m in msgs:
                        self._log(m)
                    if success:
                        self._log(f"  → 保存: {out_path}")

            except Exception as e:
                self._log(f"  エラー: {e}")
                traceback.print_exc()

        if merge and all_gdfs:
            self._log(f"\n--- 統合GeoPackage書き出し ---")
            success, msgs = save_to_geopackage(
                all_gdfs, output, source_epsg, target_epsg,
                split_by_layer=self.split_by_layer.get(),
            )
            for m in msgs:
                self._log(m)
            if success:
                self._log(f"→ 保存: {output}")

        self._log("\n変換完了！")

    def _convert_sfc(self, filepath: str) -> dict[str, gpd.GeoDataFrame]:
        """SFCファイルを変換"""
        parser = SfcParser()
        success = parser.parse_file(filepath)

        if parser.warnings:
            for w in parser.warnings[:5]:
                self._log(f"  ⚠ {w}")
            if len(parser.warnings) > 5:
                self._log(f"  ... 他 {len(parser.warnings) - 5} 件の警告")

        if not success:
            self._log("  SFCパースに失敗しました")
            return {}

        self._log(f"  レイヤ: {len(parser.layers)}, フィーチャ: {len(parser.features)}")
        return parser.to_geodataframes()

    def _get_scale_denominator(self) -> int:
        """縮尺分母を取得"""
        try:
            val = self.scale_denominator.get().split(' ')[0].strip()
            return int(val)
        except (ValueError, IndexError):
            return 0

    def _get_quality(self) -> int:
        """品質値を取得"""
        try:
            return int(self.quality.get())
        except ValueError:
            return 100

    def _convert_dxf(self, filepath: str) -> dict[str, gpd.GeoDataFrame]:
        """DXFファイルを変換"""
        scale_denom = self._get_scale_denominator()
        quality = self._get_quality()

        converter = DxfConverter()
        gdfs = converter.convert(
            filepath,
            scale_denominator=scale_denom,
            quality=quality,
            auto_georef=self.auto_georef.get(),
        )

        if converter.warnings:
            for w in converter.warnings:
                self._log(f"  ⚠ {w}")

        # ジオリファレンスを無効にする場合、スケーリングのみ結果を返す
        # （auto_georef=Falseの場合でもconvert内でスケーリングは適用済み）
        # auto_georef=Falseの場合、方眼線アフィンをスキップさせるため
        # convertメソッドに渡す必要がある → 後述

        for key, gdf in gdfs.items():
            if not gdf.empty:
                self._log(f"  {key}: {len(gdf)} フィーチャ")

        return gdfs


# ============================================================
# CLI Mode
# ============================================================
def cli_main():
    """コマンドライン実行"""
    import argparse
    parser = argparse.ArgumentParser(description="SFC/DXF → GeoPackage Converter")
    parser.add_argument('input', nargs='+', help='入力ファイル (*.sfc, *.dxf)')
    parser.add_argument('-o', '--output', required=True, help='出力GeoPackageパス')
    parser.add_argument('-z', '--zone', type=int, default=9, help='平面直角座標系の系番号 (1-19)')
    parser.add_argument('-d', '--datum', default='JGD2011', choices=['JGD2011', 'JGD2000'])
    parser.add_argument('-t', '--target-epsg', type=int, default=None, help='出力EPSG (省略でソースと同じ)')
    parser.add_argument('-s', '--scale', type=int, default=0,
                        help='図面縮尺分母 (例: 300 = 1/300図面, 0=スケーリング無し)')
    parser.add_argument('-q', '--quality', type=int, default=100,
                        choices=[100, 50, 30], help='出力品質 (100=高精細, 50=標準, 30=軽量)')
    parser.add_argument('--no-georef', action='store_true',
                        help='方眼線からの自動ジオリファレンスを無効化')
    parser.add_argument('--split-layers', action='store_true', default=True,
                        help='DXFレイヤ構造を保持して分割保存 (デフォルト: 有効)')
    parser.add_argument('--no-split', action='store_true',
                        help='レイヤ分割を無効化')
    args = parser.parse_args()

    if args.datum == 'JGD2011':
        source_epsg = JGD2011_EPSG.get(args.zone, 6677)
    else:
        source_epsg = JGD2000_EPSG.get(args.zone, 2451)

    split_by_layer = not args.no_split

    print(f"ソースCRS: EPSG:{source_epsg}")
    if args.scale > 0:
        print(f"図面縮尺: 1/{args.scale}")
    print(f"品質: {args.quality}%")
    print(f"自動ジオリファレンス: {'OFF' if args.no_georef else 'ON'}")
    print(f"レイヤ分割: {'ON' if split_by_layer else 'OFF'}")

    all_gdfs = {}
    for filepath in args.input:
        ext = Path(filepath).suffix.lower()
        print(f"\n処理中: {filepath}")

        if ext == '.sfc':
            p = SfcParser()
            p.parse_file(filepath)
            gdfs = p.to_geodataframes()
        elif ext == '.dxf':
            c = DxfConverter()
            gdfs = c.convert(
                filepath,
                scale_denominator=args.scale,
                quality=args.quality,
                auto_georef=not args.no_georef,
            )
            for w in c.warnings:
                print(f"  ⚠ {w}")
        else:
            print(f"  未対応: {ext}")
            continue

        stem = Path(filepath).stem
        for lname, gdf in gdfs.items():
            key = f"{stem}_{lname}"
            all_gdfs[key] = gdf
            print(f"  {lname}: {len(gdf)} features")

    if all_gdfs:
        success, msgs = save_to_geopackage(
            all_gdfs, args.output, source_epsg, args.target_epsg,
            split_by_layer=split_by_layer,
        )
        for m in msgs:
            print(m)


# ============================================================
# Entry Point
# ============================================================
def main():
    if len(sys.argv) > 1 and not sys.argv[1].startswith('--gui'):
        cli_main()
    else:
        root = tk.Tk()
        app = ConverterApp(root)
        root.mainloop()


if __name__ == '__main__':
    main()
