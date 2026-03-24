"""MZ0000_FlowchartTool 図形配置モジュール。

Powered by Auto (Cursor) (rev014)
Rule 2.1.5: 認知負荷の管理に基づき、excel_engine.pyから図形配置ロジックを分離。
"""
import logging
from typing import Any, Dict, List, Tuple
import pywintypes
from app.constants import ExcelConstants, FONT_FAMILY

logger = logging.getLogger("MZ0000")


def place_shapes(sheet: Any, row_map: Dict, row_heights: Dict, 
                 base_left: float, base_top: float, w_fix: float, 
                 gv: float, gh: float, theme: Dict, stop_event: Any) -> Tuple:
    """図形を適切な位置に配置し、スタイルを設定する。
    
    Args:
        sheet (Any): Excelワークシートオブジェクト。
        row_map (Dict): 行インデックスをキーとするノードマップ。
        row_heights (Dict): 行インデックスをキーとする高さの辞書。
        base_left (float): 基準となる左端位置。
        base_top (float): 基準となる上端位置。
        w_fix (float): 図形の固定幅。
        gv (float): 垂直方向の間隔。
        gh (float): 水平方向の間隔。
        theme (Dict): テーマ設定。
        stop_event (Any): 処理中止を監視するためのイベント。
        
    Returns:
        Tuple: (shape_map, standalone_names, diamond_info, bounds) の4要素タプル。
            - shape_map: ノードIDをキーとする図形オブジェクトの辞書。
            - standalone_names: 単独図形の名前リスト。
            - diamond_info: 判断図形の情報リスト。
            - bounds: (left, top, right, bottom) の境界座標タプル。
    """
    shape_map = {}
    standalone_names = []
    diamond_info = []
    l_list, t_list, r_list, b_list = [], [], [], []
    
    current_top = base_top
    last_ri = -1
    
    for ri in sorted(row_map.keys()):
        if last_ri != -1:
            current_top += row_heights[last_ri] + gv
        
        for n in row_map[ri]:
            if stop_event.is_set():
                break
            
            left_pos = base_left + (n["level"] * (w_fix + gh))
            stype = n["type"]
            
            # 図形種別判定
            stype_code = ExcelConstants.MSOSHAPE_RECTANGLE
            is_diamond = False
            if "判断" in stype:
                stype_code = ExcelConstants.MSOSHAPE_DIAMOND
                is_diamond = True
            elif any(x in stype for x in ["端子", "開始", "終了"]):
                stype_code = ExcelConstants.MSOSHAPE_ROUNDED_RECTANGLE
            elif any(x in stype for x in ["入出力", "データ"]):
                stype_code = ExcelConstants.MSOSHAPE_PARALLELOGRAM
            elif "手動入力" in stype:
                stype_code = ExcelConstants.MSOSHAPE_MANUAL_INPUT
            
            # 配置
            shp_h = row_heights[ri] * 1.3 if is_diamond else row_heights[ri]
            v_off = (shp_h - row_heights[ri]) / 2 if is_diamond else 0
            
            shp = sheet.Shapes.AddShape(stype_code, left_pos, current_top - v_off, w_fix, shp_h)
            shp.Fill.ForeColor.RGB = 0xFFFFFF
            shp.Line.ForeColor.RGB = theme["shape_line"]
            
            if is_diamond:
                diamond_info.append({
                    "shp": shp,
                    "l": left_pos,
                    "t": current_top,
                    "h": row_heights[ri],
                    "txt": n["full_text"]
                })
            else:
                standalone_names.append(shp.Name)
                set_text_style(shp, n["full_text"], is_manual=(stype_code == ExcelConstants.MSOSHAPE_MANUAL_INPUT))
            
            shape_map[n["id"]] = shp
            l_list.append(float(shp.Left))
            t_list.append(float(shp.Top))
            r_list.append(float(shp.Left) + float(shp.Width))
            b_list.append(float(shp.Top) + float(shp.Height))
        
        last_ri = ri
        if stop_event.is_set():
            break
        
    bounds = (min(l_list), min(t_list), max(r_list), max(b_list)) if l_list else (0, 0, 0, 0)
    return shape_map, standalone_names, diamond_info, bounds


def set_text_style(shp: Any, text: str, is_manual: bool = False) -> None:
    """図形のテキストスタイルを一括設定する。
    
    Args:
        shp (Any): Excel図形オブジェクト。
        text (str): 設定するテキスト内容。
        is_manual (bool): 手動入力図形の場合True。デフォルトはFalse。
    """
    try:
        tf2 = shp.TextFrame2
        tf2.TextRange.Text = text
        tf2.TextRange.ParagraphFormat.Alignment = ExcelConstants.MSO_ALIGN_CENTER
        tf2.VerticalAnchor = ExcelConstants.MSO_ANCHOR_MIDDLE
        tf2.WordWrap = True
        tf2.TextRange.Font.Size = 11
        tf2.TextRange.Font.Name = FONT_FAMILY
        tf2.TextRange.Font.Fill.ForeColor.RGB = 0
        
        if is_manual:
            shp.Adjustments.Item(1, 0.2)
            shp.TextFrame.MarginTop = 12.0
        else:
            shp.TextFrame.MarginLeft = shp.TextFrame.MarginRight = 0
            shp.TextFrame.MarginTop = shp.TextFrame.MarginBottom = 0
    except (pywintypes.com_error, AttributeError) as e:
        logger.error(f"text_style_error | error={e}")
