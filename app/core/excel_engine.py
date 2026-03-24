"""MZ0000_FlowchartTool Excel操作・描画エンジン。

Powered by Auto (Cursor) (rev014)
Rule 2.1: SoC (Separation of Concerns) に基づき、UIに依存しない純粋なロジックを保持する。
Rule 2.1.5: 認知負荷の管理に基づき、図形配置・コネクタ接続・グループ化ロジックを分離モジュールに移行。
"""
import logging
import threading
from typing import Optional, Any, List, Dict, Tuple
import pythoncom
import win32com.client
import pywintypes
from app.constants import ExcelConstants
from app.core.shape_placer import place_shapes
from app.core.connector_manager import connect_nodes
from app.core.group_manager import finalize_composites, add_frame_and_title, create_final_groups

logger = logging.getLogger("MZ0000")


def get_excel_app() -> Optional[Any]:
    """Excelアプリケーションオブジェクトを安全に取得する。
    
    Returns:
        Optional[Any]: Excelアプリケーションオブジェクト。取得できない場合はNone。
    """
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except (pywintypes.com_error, AttributeError):
        return None


class ExcelFlowchartEngine:
    """Excel上での描画ロジックをカプセル化したエンジンクラス。"""

    def __init__(self, stop_event: threading.Event) -> None:
        """初期化。
        
        Args:
            stop_event (threading.Event): 処理中止を監視するためのイベント。
        """
        self.stop_event = stop_event
        self.last_group_name: Optional[str] = None

    def draw(self, 
             is_full_mode: bool, 
             config: Dict[str, Any], 
             theme: Dict[str, Any]) -> str:
        """フローチャートを描画するメインエントリポイント。
        
        Args:
            is_full_mode (bool): 表全体から生成する場合True、選択範囲のみの場合False。
            config (Dict[str, Any]): 描画設定（height, width, gap_v, gap_hを含む）。
            theme (Dict[str, Any]): テーマ設定（connector, shape_line, labelを含む）。
            
        Returns:
            str: 作成したグループ図形の名前。失敗時は空文字列。
            
        Raises:
            RuntimeError: Excelが起動していない場合。
        """
        app = get_excel_app()
        if not app:
            logger.error("excel_not_found | Excel is not running.")
            raise RuntimeError("Excelが起動していません。")

        # 実行前のサイレント化
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        
        try:
            sel = app.Selection
            r_tgt = sel.CurrentRegion if is_full_mode else sel
            data = r_tgt.Value
            
            if not data or not isinstance(data, tuple):
                logger.warning("no_data_selected | Selection is empty or invalid.")
                return ""

            sheet = app.ActiveSheet
            start_cell = r_tgt.Cells(1, 1)
            
            # タイトルの取得
            title_txt = "フローチャート"
            if is_full_mode:
                for i in range(-5, 1):
                    row_idx = max(1, start_cell.Row + i)
                    c = sheet.Cells(row_idx, start_cell.Column)
                    if c.Interior.Color == ExcelConstants.TITLE_BG_COLOR and c.Value:
                        title_txt = str(c.Value)
                        break

            # 描画パラメータ
            base_left = float(start_cell.Left)
            base_top = float(start_cell.Top)
            h_min = float(config["height"])
            w_fix = float(config["width"])
            gv = float(config["gap_v"])
            gh = float(config["gap_h"])

            # 1. データ解析
            nodes, row_map = self._parse_data(data)
            if not nodes:
                return ""

            # 2. 高さ計算 (テキスト量に応じた動的調整)
            row_heights = self._calculate_row_heights(sheet, row_map, w_fix, h_min)

            # 3. 図形配置
            shape_map, standalone_names, diamond_info, bounds = place_shapes(
                sheet, row_map, row_heights, base_left, base_top, w_fix, gv, gh, theme, self.stop_event
            )

            if self.stop_event.is_set(): 
                logger.info("draw_cancelled_before_connect")
                return ""

            # 4. コネクタ接続
            connector_names = connect_nodes(sheet, nodes, shape_map, theme, self.stop_event)

            if self.stop_event.is_set(): 
                logger.info("draw_cancelled_before_finalize")
                return ""

            # 5. コンポジット（判断図形）の仕上げ
            composite_pairs = finalize_composites(sheet, diamond_info, w_fix)

            # 6. 外枠とタイトル (フルモード時)
            extra_names = []
            if is_full_mode:
                extra_names = add_frame_and_title(sheet, bounds, title_txt)

            # 7. グループ化
            all_names = standalone_names + connector_names + extra_names
            group_name = create_final_groups(sheet, all_names, composite_pairs)
            
            logger.info(f"draw_completed | group_name={group_name}")
            return group_name

        finally:
            app.ScreenUpdating = True
            app.DisplayAlerts = True

    def _parse_data(self, data: Tuple) -> Tuple[List[Dict], Dict[int, List[Dict]]]:
        """Excelデータを内部ノード構造に変換する。
        
        Args:
            data (Tuple): Excelから取得したセルデータのタプル。
            
        Returns:
            Tuple[List[Dict], Dict[int, List[Dict]]]: 
                - ノードリスト（各ノードはid, type, full_text, dests_down, dests_right, level, ridxを含む）。
                - 行インデックスをキーとするノードマップ。
        """
        nodes = []
        row_map = {}
        col_count = len(data[0]) if data else 0
        
        def norm(v: Any) -> str:
            return str(v).split('.')[0].strip() if v is not None and v != "" else ""

        for i, row in enumerate(data):
            nid = norm(row[0])
            if not nid.isdigit():
                continue
            
            # 8列 or 7列 の判定
            if col_count >= 8:
                txts = [str(row[j]) for j in range(5, min(8, len(row))) if row[j]]
                d_down = [norm(d) for d in str(row[2]).split(',') if norm(d)]
                d_right = [norm(d) for d in str(row[3]).split(',') if norm(d)]
                level = row[4]
            elif col_count == 7:
                txts = [str(row[j]) for j in range(4, min(7, len(row))) if row[j]]
                d_down = [norm(d) for d in str(row[2]).split(',') if norm(d)]
                d_right = []
                level = row[3]
            else:
                txts = [str(row[2])] if len(row) > 2 and row[2] else []
                d_down = [norm(d) for d in str(row[3]).split(',') if norm(d)]
                d_right = []
                level = row[4] if len(row) > 4 else 0

            node = {
                "id": nid,
                "type": str(row[1]) if row[1] else "処理",
                "full_text": "\n".join(txts),
                "dests_down": d_down,
                "dests_right": d_right,
                "level": int(float(level)) if level is not None else 0,
                "ridx": i
            }
            nodes.append(node)
            row_map.setdefault(i, []).append(node)
        return nodes, row_map

    def _calculate_row_heights(self, sheet: Any, row_map: Dict, w_fix: float, h_min: float) -> Dict[int, float]:
        """テキスト量に基づき各行の高さを動的に計算する。
        
        Args:
            sheet (Any): Excelワークシートオブジェクト。
            row_map (Dict): 行インデックスをキーとするノードマップ。
            w_fix (float): 図形の固定幅。
            h_min (float): 図形の最小高さ。
            
        Returns:
            Dict[int, float]: 行インデックスをキーとする高さの辞書。
        """
        heights = {}
        # 一時図形の生成（清算義務 Layer 2.1.3）
        temp_shp = sheet.Shapes.AddShape(ExcelConstants.MSOSHAPE_RECTANGLE, -5000, -5000, w_fix, h_min)
        try:
            for ri, r_nodes in row_map.items():
                max_h = h_min
                for n in r_nodes:
                    temp_shp.TextFrame2.TextRange.Text = n["full_text"]
                    temp_shp.TextFrame2.AutoSize = 1
                    max_h = max(max_h, float(temp_shp.Height) + 15.0)
                heights[ri] = max_h
        finally:
            temp_shp.Delete()
        return heights
