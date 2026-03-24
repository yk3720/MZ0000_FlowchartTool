"""MZ0000_FlowchartTool コネクタ接続管理モジュール。

Powered by Auto (Cursor) (rev014)
Rule 2.1.5: 認知負荷の管理に基づき、excel_engine.pyからコネクタ接続ロジックを分離。
"""
import logging
from typing import Any, Dict, List, Optional
import pywintypes
from app.constants import ExcelConstants, FONT_FAMILY

logger = logging.getLogger("MZ0000")


def connect_nodes(sheet: Any, nodes: List[Dict], shape_map: Dict, theme: Dict, stop_event: Any) -> List[str]:
    """ノード間をコネクタで接続し、ラベル（Yes/No）を配置する。
    
    Args:
        sheet (Any): Excelワークシートオブジェクト。
        nodes (List[Dict]): ノードリスト。
        shape_map (Dict): ノードIDをキーとする図形オブジェクトの辞書。
        theme (Dict): テーマ設定。
        stop_event (Any): 処理中止を監視するためのイベント。
        
    Returns:
        List[str]: 作成したコネクタとラベルの名前リスト。
    """
    names = []
    node_dict = {n["id"]: n for n in nodes}
    
    for n in nodes:
        s_shp = shape_map.get(n["id"])
        if not s_shp or stop_event.is_set():
            continue
        
        for direction, dests in [("down", n["dests_down"]), ("right", n["dests_right"])]:
            for did in dests:
                t_shp = shape_map.get(did)
                t_node = node_dict.get(did)
                if not t_shp or not t_node:
                    continue
                
                is_loop = t_node["ridx"] < n["ridx"]
                level_diff = t_node["level"] - n["level"]
                
                # 接続ロジック
                s_site, t_site = ExcelConstants.CONNECTOR_SITE_BOTTOM, ExcelConstants.CONNECTOR_SITE_TOP
                c_type = ExcelConstants.MSOCONNECTOR_STRAIGHT
                
                if direction == "down":
                    if level_diff != 0 or is_loop:
                        c_type = ExcelConstants.MSOCONNECTOR_ELBOW
                        if level_diff < 0:
                            t_site = ExcelConstants.CONNECTOR_SITE_LEFT
                        elif level_diff > 0:
                            s_site, t_site = ExcelConstants.CONNECTOR_SITE_RIGHT, ExcelConstants.CONNECTOR_SITE_TOP
                else:
                    s_site, t_site = ExcelConstants.CONNECTOR_SITE_RIGHT, ExcelConstants.CONNECTOR_SITE_TOP
                    c_type = ExcelConstants.MSOCONNECTOR_ELBOW
                    if level_diff == 0 and is_loop:
                        t_site = ExcelConstants.CONNECTOR_SITE_RIGHT
                    elif level_diff < 0:
                        t_site = ExcelConstants.CONNECTOR_SITE_LEFT

                # 同一列かつ順方向なら直線
                if direction == "down" and abs(s_shp.Left - t_shp.Left) < 5 and not is_loop:
                    c_type = ExcelConstants.MSOCONNECTOR_STRAIGHT
                    
                try:
                    conn = sheet.Shapes.AddConnector(c_type, 0, 0, 10, 10)
                    names.append(conn.Name)
                    conn.ConnectorFormat.BeginConnect(s_shp, s_site)
                    conn.ConnectorFormat.EndConnect(t_shp, t_site)
                    conn.Line.ForeColor.RGB = theme["connector"]
                    conn.Line.Weight = 2.25
                    conn.Line.EndArrowheadStyle = 3
                    
                    # 判断ラベル
                    if "判断" in n["type"]:
                        lbl_name = add_decision_label(sheet, s_shp, direction)
                        if lbl_name:
                            names.append(lbl_name)
                except (pywintypes.com_error, AttributeError) as e:
                    logger.error(f"connector_error | from={n['id']} | to={did} | error={e}")
    return names


def add_decision_label(sheet: Any, s_shp: Any, direction: str) -> Optional[str]:
    """判断図形からの分岐に Yes/No ラベルを配置する。
    
    Args:
        sheet (Any): Excelワークシートオブジェクト。
        s_shp (Any): 判断図形オブジェクト。
        direction (str): 分岐方向（"right" または "down"）。
        
    Returns:
        Optional[str]: 作成したラベル図形の名前。失敗時はNone。
    """
    try:
        lw, lh = 30, 16
        if direction == "right":
            lx = s_shp.Left + s_shp.Width - 10
            ly = s_shp.Top + (s_shp.Height / 2) - lh - 3
            txt = "No"
        else:
            lx = s_shp.Left + (s_shp.Width / 2) - lw - 3
            ly = s_shp.Top + s_shp.Height + 3
            txt = "Yes"
        
        lbl = sheet.Shapes.AddTextbox(1, lx, ly, lw, lh)
        lbl.Fill.Visible = False
        lbl.Line.Visible = False
        tr = lbl.TextFrame2.TextRange
        tr.Text = txt
        tr.Font.Size = 9
        tr.Font.Bold = True
        tr.Font.Name = FONT_FAMILY
        tr.Font.Fill.ForeColor.RGB = 0
        tr.ParagraphFormat.Alignment = ExcelConstants.MSO_ALIGN_CENTER
        lbl.TextFrame2.VerticalAnchor = ExcelConstants.MSO_ANCHOR_MIDDLE
        return lbl.Name
    except (pywintypes.com_error, AttributeError):
        logger.warning(f"decision_label_creation_failed | direction={direction}")
        return None
