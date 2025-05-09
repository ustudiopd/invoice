from PyQt5.QtWidgets import QTableWidget
from PyQt5.QtCore import Qt


class ZoomableTableWidget(QTableWidget):
    """확대/축소가 가능한 테이블 위젯"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.zoom_factor = 1.0
        self.min_zoom = 0.1  # 10%
        self.max_zoom = 2.5  # 250%
        self.zoom_step = 0.1
        self.base_font_size = 9
        self.base_row_height = 20
        self.last_mouse_pos = None
        self.snap_zoom_levels = [0.9, 1.0, 1.25, 1.5, 1.75, 2.0, 2.5]
        self.snap_threshold = 0.05
        # 초기 열 너비/행 높이 저장
        self._init_col_widths = []
        self._init_row_heights = []

    def set_initial_sizes(self):
        """초기 열 너비/행 높이 저장 (엑셀 로드 후 1회 호출)"""
        self._init_col_widths = [
            self.columnWidth(col)
            for col in range(self.columnCount())
        ]
        self._init_row_heights = [
            self.rowHeight(row)
            for row in range(self.rowCount())
        ]

    def wheelEvent(self, event):
        """Ctrl + 휠로 확대/축소 (방향 전환 즉시 반응)"""
        if event.modifiers() == Qt.ControlModifier:
            cursor_pos = event.pos()
            self.last_mouse_pos = cursor_pos
            viewport_rect = self.viewport().rect()
            cursor_rel_x = cursor_pos.x() / viewport_rect.width()
            cursor_rel_y = cursor_pos.y() / viewport_rect.height()
            h_scroll = self.horizontalScrollBar()
            v_scroll = self.verticalScrollBar()
            delta = event.angleDelta().y()
            # 휠 방향에 따라 확대/축소
            if delta > 0:
                new_zoom = min(
                    self.zoom_factor + self.zoom_step,
                    self.max_zoom
                )
            elif delta < 0:
                new_zoom = max(
                    self.zoom_factor - self.zoom_step,
                    self.min_zoom
                )
            else:
                new_zoom = self.zoom_factor
            # 스냅 기능 적용
            new_zoom = self._apply_snap_zoom(new_zoom)
            # 무조건 zoom_factor 갱신 및 적용
            self.zoom_factor = new_zoom
            self._apply_zoom()
            # 스크롤 위치 조정
            new_h_max = h_scroll.maximum()
            new_v_max = v_scroll.maximum()
            new_h_value = int(new_h_max * cursor_rel_x)
            new_v_value = int(new_v_max * cursor_rel_y)
            h_scroll.setValue(new_h_value)
            v_scroll.setValue(new_v_value)
            event.accept()
        else:
            super().wheelEvent(event)

    def _apply_snap_zoom(self, zoom_value):
        """가장 가까운 스냅 레벨로 확대/축소 값 조정"""
        closest_snap = None
        min_diff = float('inf')
        
        for snap_level in self.snap_zoom_levels:
            diff = abs(zoom_value - snap_level)
            if diff < min_diff:
                min_diff = diff
                closest_snap = snap_level
        
        if min_diff < self.snap_threshold:
            return closest_snap
        return zoom_value

    def set_zoom(self, value):
        """슬라이더로부터 확대/축소 값 설정"""
        # 현재 뷰포트의 중심점 계산
        viewport_rect = self.viewport().rect()
        center_x = viewport_rect.width() / 2
        center_y = viewport_rect.height() / 2
        
        # 현재 스크롤 위치 저장
        h_scroll = self.horizontalScrollBar()
        v_scroll = self.verticalScrollBar()
        
        # 새로운 확대/축소 값 계산
        zoom_value = value / 100.0
        new_zoom = self._apply_snap_zoom(zoom_value)
        
        # 확대/축소 비율이 변경된 경우에만 처리
        if new_zoom != self.zoom_factor:
            self.zoom_factor = new_zoom
            
            # 확대/축소 적용
            self._apply_zoom()
            
            # 새로운 스크롤 위치 계산
            new_h_max = h_scroll.maximum()
            new_v_max = v_scroll.maximum()
            
            # 중심점 기준으로 스크롤 위치 조정
            new_h_value = int(new_h_max * (center_x / viewport_rect.width()))
            new_v_value = int(new_v_max * (center_y / viewport_rect.height()))
            
            # 스크롤 위치 업데이트
            h_scroll.setValue(new_h_value)
            v_scroll.setValue(new_v_value)

    def _apply_zoom(self):
        """현재 zoom_factor를 기준값에 곱해서 적용"""
        # 폰트 크기
        font = self.font()
        font.setPointSizeF(self.base_font_size * self.zoom_factor)
        self.setFont(font)
        # 행 높이
        for row in range(self.rowCount()):
            base_h = self.base_row_height
            if self._init_row_heights and row < len(self._init_row_heights):
                base_h = self._init_row_heights[row]
            self.setRowHeight(row, int(base_h * self.zoom_factor))
        # 열 너비
        for col in range(self.columnCount()):
            base_w = 80
            if self._init_col_widths and col < len(self._init_col_widths):
                base_w = self._init_col_widths[col]
            self.setColumnWidth(col, int(base_w * self.zoom_factor)) 