from PyQt5.QtWidgets import QStyledItemDelegate
from PyQt5.QtGui import QPen, QColor
from PyQt5.QtCore import Qt

class BorderDelegate(QStyledItemDelegate):
    """openpyxl border 정보에 따라 셀 테두리만 그려주는 Delegate"""
    def paint(self, painter, option, index):
        # 1) 기본 렌더링 (배경·글자 등)
        super().paint(painter, option, index)
        # 2) border_info에 따라 테두리만 그림
        border_info = index.data(Qt.UserRole)
        if not isinstance(border_info, dict):
            return
        pen = QPen(QColor('#555555'))
        pen.setWidth(1)
        painter.save()
        painter.setPen(pen)
        r = option.rect
        # 위쪽 테두리
        if border_info.get("top"):
            painter.drawLine(r.topLeft(), r.topRight())
        # 아래쪽
        if border_info.get("bottom"):
            painter.drawLine(r.bottomLeft(), r.bottomRight())
        # 왼쪽
        if border_info.get("left"):
            painter.drawLine(r.topLeft(), r.bottomLeft())
        # 오른쪽
        if border_info.get("right"):
            painter.drawLine(r.topRight(), r.bottomRight())
        painter.restore() 