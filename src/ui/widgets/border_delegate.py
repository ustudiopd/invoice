from PyQt5.QtWidgets import QStyledItemDelegate
from PyQt5.QtGui import QPainter, QPen
from PyQt5.QtCore import Qt

class BorderDelegate(QStyledItemDelegate):
    """openpyxl border 정보에 따라 셀 테두리만 그려주는 Delegate"""
    def paint(self, painter, option, index):
        # 1) 기본 렌더링 (배경·글자 등)
        super().paint(painter, option, index)
        # 2) border_info에 따라 테두리만 그림
        border_info = index.data(Qt.UserRole)
        if not border_info:
            return
        painter.save()
        pen = QPen(Qt.black, 1)
        painter.setPen(pen)
        rect = option.rect
        # 위쪽 테두리
        if border_info.get("top"):
            painter.drawLine(rect.topLeft(), rect.topRight())
        # 아래쪽
        if border_info.get("bottom"):
            painter.drawLine(rect.bottomLeft(), rect.bottomRight())
        # 왼쪽
        if border_info.get("left"):
            painter.drawLine(rect.topLeft(), rect.bottomLeft())
        # 오른쪽
        if border_info.get("right"):
            painter.drawLine(rect.topRight(), rect.bottomRight())
        painter.restore() 