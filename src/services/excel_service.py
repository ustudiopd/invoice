import os
import json
import openpyxl
from PyQt5.QtWidgets import QTableWidgetItem
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtCore import Qt
from ..utils.color_utils import apply_tint


class ExcelService:
    def __init__(self):
        self.accent_colors = {1: '3B4E87'}  # accent1 파랑 계열 RGB를 하드코딩

    def load_excel(self, path, table_widget):
        """엑셀 파일을 로드하여 테이블 위젯에 표시합니다."""
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows, cols = ws.max_row, ws.max_column
        
        # 테이블 초기화
        table_widget.blockSignals(True)
        table_widget.clear()
        table_widget.setRowCount(rows)
        table_widget.setColumnCount(cols)

        # 헤더 정보 추출
        header_row, header_map = self._extract_header_info(ws)
        
        # 품목 정보 추출
        items = self._extract_items(ws, header_row, header_map)

        # 셀 스타일 적용
        self._apply_cell_styles(ws, table_widget)

        # 병합 셀 처리
        self._handle_merged_cells(ws, table_widget)

        # 열 너비/행 높이 설정
        self._set_dimensions(ws, table_widget)

        table_widget.blockSignals(False)

        # JSON 파일 저장
        json_path = os.path.splitext(path)[0] + ".json"
        self._save_to_json(json_path, ws, items)

        return json_path, items

    def _extract_header_info(self, ws):
        """헤더 정보를 추출합니다."""
        header_keywords = [
            "item", "품목", "항목", "품명",
            "상세 내역", "상 세 내 역", "상세내역",
            "description"
        ]
        quantity_keywords = ["quantity", "수량", "quant", "qty"]
        day_keywords = ["day", "일수"]
        unit_cost_keywords = [
            "unit cost", "단가", "unit krw", "unit"
        ]
        total_amount_keywords = [
            "total amount", "금액", "합계", "amount"
        ]
        won_keywords = ["won", "krw", "금액", "합계"]

        header_row = None
        header_map = {}
        amount_candidates = []

        for r in range(1, ws.max_row + 1):
            row_values = [
                str(ws.cell(row=r, column=c).value).strip().lower()
                if ws.cell(row=r, column=c).value is not None else ""
                for c in range(1, ws.max_column + 1)
            ]
            
            for idx, val in enumerate(row_values):
                if any(k in val for k in header_keywords):
                    if "상세" in val:
                        header_map["Description"] = idx
                    else:
                        header_map["Item"] = idx
                if any(k in val for k in quantity_keywords):
                    header_map["Quantity"] = idx
                if any(k in val for k in day_keywords):
                    header_map["day"] = idx
                if any(k in val for k in unit_cost_keywords):
                    header_map["Unit Cost"] = idx
                if any(k in val for k in total_amount_keywords):
                    amount_candidates.append((idx, val))

            if amount_candidates:
                for idx, val in amount_candidates:
                    if any(w in val for w in won_keywords):
                        header_map["Total Amount"] = idx
                        break
                else:
                    header_map["Total Amount"] = amount_candidates[0][0]

            header_fields = [
                "Item", "Description", "Quantity",
                "Unit Cost", "Total Amount"
            ]
            match_count = sum(1 for f in header_fields if f in header_map)
            if match_count >= 3:
                header_row = r
                break

        return header_row, header_map

    def _extract_items(self, ws, header_row, header_map):
        """품목 정보를 추출합니다."""
        items = []
        if header_row:
            for r in range(header_row + 1, ws.max_row + 1):
                row_texts = [
                    str(ws.cell(row=r, column=c+1).value).strip()
                    for c in range(ws.max_column)
                    if (
                        ws.cell(row=r, column=c+1).value is not None and
                        str(ws.cell(row=r, column=c+1).value).strip() != ""
                    )
                ]
                item_name = " | ".join(row_texts) if row_texts else None
                
                if (
                    not item_name or
                    str(item_name).strip() == "" or
                    len(str(item_name).strip()) <= 2 or
                    any(x in str(item_name) for x in [
                        "합계", "총액", "vat", "참고"
                    ])
                ):
                    continue

                quantity = ws.cell(
                    row=r, column=header_map.get("Quantity") + 1
                ).value if header_map.get("Quantity") is not None else None
                
                day = ws.cell(
                    row=r, column=header_map.get("day") + 1
                ).value if header_map.get("day") is not None else None
                
                unit_cost = ws.cell(
                    row=r, column=header_map.get("Unit Cost") + 1
                ).value if header_map.get("Unit Cost") is not None else None
                
                total_amount = ws.cell(
                    row=r, column=header_map.get("Total Amount") + 1
                ).value if header_map.get("Total Amount") is not None else None

                item = {
                    "description": str(item_name),
                    "quantity": quantity,
                    "day": day,
                    "unit_price": unit_cost,
                    "amount": total_amount
                }
                items.append(item)

        return items

    def _apply_cell_styles(self, ws, table_widget):
        """셀 스타일을 적용합니다."""
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                text = cell.value or ""
                item = QTableWidgetItem(str(text))

                # 폰트
                f = cell.font
                point_size = int(f.sz) if f.sz is not None else -1
                qf = QFont(f.name, point_size)
                qf.setBold(f.b)
                qf.setItalic(f.i)
                item.setFont(qf)

                # 배경색
                color = self._get_cell_color(cell)
                if (color and 
                    (color.red(), color.green(), color.blue()) != (255, 255, 255)):
                    item.setBackground(color)

                # 정렬
                align = cell.alignment
                qt_align = 0
                if align.horizontal == 'center':
                    qt_align |= Qt.AlignHCenter
                elif align.horizontal == 'right':
                    qt_align |= Qt.AlignRight
                else:
                    qt_align |= Qt.AlignLeft
                if align.vertical == 'center':
                    qt_align |= Qt.AlignVCenter
                elif align.vertical == 'bottom':
                    qt_align |= Qt.AlignBottom
                else:
                    qt_align |= Qt.AlignTop
                item.setTextAlignment(qt_align)

                # 테두리
                b = cell.border
                border_info = {
                    "top": bool(b.top and b.top.style),
                    "bottom": bool(b.bottom and b.bottom.style),
                    "left": bool(b.left and b.left.style),
                    "right": bool(b.right and b.right.style),
                }
                item.setData(Qt.UserRole, border_info)
                table_widget.setItem(r-1, c-1, item)

    def _get_cell_color(self, cell):
        """셀의 배경색을 가져옵니다."""
        try:
            fill = cell.fill
            fg = fill.fgColor if hasattr(fill, 'fgColor') else None
            
            if (fill.patternType in ('solid', 'gray125', 'darkGrid', 'lightGrid')
                    and fg):
                # 1) Theme 컬러 + 틴트
                if fg.type == 'theme':
                    tint = getattr(fg, 'tint', 0.0)
                    hex_rgb = self.accent_colors[1]
                    return apply_tint(hex_rgb, tint)
                # 2) RGB 컬러
                elif fg.type == 'rgb' and fg.rgb:
                    rgb = fg.rgb[2:] if fg.rgb.startswith('FF') else fg.rgb
                    return QColor(
                        int(rgb[0:2], 16),
                        int(rgb[2:4], 16),
                        int(rgb[4:6], 16)
                    )
                # 3) Indexed 컬러
                elif fg.type == 'indexed' and fg.indexed is not None:
                    from openpyxl.styles.colors import COLOR_INDEX
                    idx = fg.indexed
                    if 0 <= idx < len(COLOR_INDEX):
                        hexcol = COLOR_INDEX[idx][2:]
                        return QColor(
                            int(hexcol[0:2], 16),
                            int(hexcol[2:4], 16),
                            int(hexcol[4:6], 16)
                        )
                # 4) Gradient Fill
                elif hasattr(fill, 'gradientType') and fill.gradientType:
                    stops = getattr(fill, 'stop', None)
                    if (stops and hasattr(stops[0], 'color')
                            and hasattr(stops[0].color, 'rgb')):
                        rgb = (stops[0].color.rgb[2:]
                               if stops[0].color.rgb.startswith('FF')
                               else stops[0].color.rgb)
                        return QColor(
                            int(rgb[0:2], 16),
                            int(rgb[2:4], 16),
                            int(rgb[4:6], 16)
                        )
        except Exception:
            pass
        return None

    def _handle_merged_cells(self, ws, table_widget):
        """병합된 셀을 처리합니다."""
        for merged in ws.merged_cells.ranges:
            r0, c0 = merged.min_row-1, merged.min_col-1
            rs = merged.max_row - merged.min_row + 1
            cs = merged.max_col - merged.min_col + 1
            table_widget.setSpan(r0, c0, rs, cs)

    def _set_dimensions(self, ws, table_widget):
        """열 너비와 행 높이를 설정합니다."""
        # 열 너비
        for idx, col_dim in ws.column_dimensions.items():
            col = openpyxl.utils.column_index_from_string(idx) - 1
            if col < ws.max_column and col_dim.width:
                table_widget.setColumnWidth(col, int(col_dim.width * 7))
        # 행 높이
        for r, row_dim in ws.row_dimensions.items():
            if row_dim.height is not None and r-1 < ws.max_row:
                table_widget.setRowHeight(r-1, int(row_dim.height * 1.2))

    def _save_to_json(self, json_path, ws, items):
        """데이터를 JSON 파일로 저장합니다."""
        data_dict = {
            "meta": {
                "file_name": os.path.basename(json_path)
            },
            "items": items,
            "header": {},
            "summary": {}
        }

        # 헤더 정보 추출
        header_labels = [
            "DATE", "QUOTATION #", "Payment date", "SHIP TO",
            "발급일", "공급자", "등록번호", "상호", "대표이사", "사업자 주소"
        ]
        for r in range(1, 16):
            for c in range(1, ws.max_column):
                raw = ws.cell(row=r, column=c).value
                if raw is None:
                    continue
                key = str(raw).strip().rstrip(":")
                if key in header_labels:
                    val = ws.cell(row=r, column=c+1).value \
                        or ws.cell(row=r+1, column=c).value
                    data_dict["header"][key] = val

        # 요약 정보 추출
        summary_labels = {
            "subtotal": ["Sub total", "소계"],
            "tax_rate": ["Tax rate", "세율"],
            "tax_due": ["Tax due", "세금", "Tax due"],
            "other": ["Other", "기타"],
            "total_due": ["TOTAL Due", "총액", "합계", "TOTAL"]
        }
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column):
                raw = ws.cell(row=r, column=c).value
                if raw is None:
                    continue
                txt = str(raw).strip()
                for fld, keys in summary_labels.items():
                    if any(k in txt for k in keys):
                        data_dict["summary"][fld] = (
                            ws.cell(row=r, column=c+1).value
                        )

        # 상단 고정 레이블 추출
        for r in range(3, 10):
            raw_key = ws.cell(row=r, column=4).value
            if raw_key:
                key = str(raw_key).strip().rstrip(':')
                data_dict["header"][key] = ws.cell(row=r, column=5).value

        # JSON 파일 저장
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data_dict, f, ensure_ascii=False, indent=2, default=str) 