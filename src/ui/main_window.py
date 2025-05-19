import json
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QFrame, QTextBrowser,
    QTextEdit, QSplitter, QLabel, QMessageBox
)
from PyQt5.QtCore import Qt
from src.config.settings import GPT_API_KEY, GPT_MODEL
from src.services.gpt_service import ask_gpt_api
from src.services.excel_service import ExcelService
from src.services.gpt_service import GPTService
from src.ui.widgets.zoomable_table import ZoomableTableWidget
from src.ui.widgets.border_delegate import BorderDelegate


class ExcelGPTViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel GPT Viewer")
        self.setGeometry(100, 100, 1200, 800)
        self.json_path = None
        self.excel_path = None
        self.excel_service = ExcelService()
        self.gpt_service = GPTService(GPT_API_KEY, GPT_MODEL)

        self._init_ui()

    def _init_ui(self):
        """UI 초기화"""
        # 메인 레이아웃
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # QSplitter(좌:엑셀, 우:챗)
        self.splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(self.splitter)

        # 1) 엑셀 뷰어 패널
        excel_panel = QWidget()
        excel_layout = QVBoxLayout(excel_panel)
        excel_layout.setContentsMargins(8, 8, 8, 8)
        
        # 상단 버튼 레이아웃
        top_layout = QHBoxLayout()
        file_btn = QPushButton("엑셀 파일 열기")
        file_btn.clicked.connect(self.open_excel)
        top_layout.addWidget(file_btn)
        excel_layout.addLayout(top_layout)
        
        # ZoomableTableWidget으로 변경
        self.excel_view = ZoomableTableWidget()
        # 기본 그리드(격자) 끄기
        self.excel_view.setShowGrid(False)
        # 셀별 테두리 그리도록 Delegate 설정
        self.excel_view.setItemDelegate(BorderDelegate(self.excel_view))
        self.excel_view.cellChanged.connect(self.on_cell_changed)
        excel_layout.addWidget(self.excel_view)
        
        self.splitter.addWidget(excel_panel)

        # 2) 챗봇 패널
        chat_frame = QFrame()
        chat_layout = QVBoxLayout(chat_frame)
        chat_layout.setContentsMargins(8, 8, 8, 8)
        chat_layout.setSpacing(8)
        self.model_label = QLabel(f"모델: {GPT_MODEL}")
        self.model_label.setStyleSheet("font-weight:bold;color:#0057b8;")
        chat_layout.addWidget(self.model_label)
        self.chat_output = QTextBrowser()
        self.chat_output.setOpenExternalLinks(True)
        self.chat_output.setMinimumHeight(180)
        chat_layout.addWidget(self.chat_output)
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        chat_layout.addWidget(line)
        self.chat_input = QTextEdit()
        self.chat_input.setPlaceholderText("질문을 입력하세요...")
        self.chat_input.setMinimumHeight(40)
        self.chat_input.setStyleSheet("border:2px solid #000;")
        # 엔터키 이벤트 추가
        self.chat_input.installEventFilter(self)
        chat_layout.addWidget(self.chat_input)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch(1)
        self.send_btn = QPushButton("질문하기")
        self.send_btn.clicked.connect(self.ask_gpt)
        btn_layout.addWidget(self.send_btn)
        btn_layout.addStretch(1)
        btn_widget = QWidget()
        btn_widget.setLayout(btn_layout)
        chat_layout.addWidget(btn_widget)
        self.splitter.addWidget(chat_frame)
        self.splitter.setSizes([900, 500])

        # 로그 메시지창 (하단)
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(80)
        style = "background:#222;color:#eee;font-size:12px;"
        self.log_output.setStyleSheet(style)
        # 메인 레이아웃에 로그창 추가 (세로로 쌓기)
        vbox = QVBoxLayout()
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(0)
        vbox.addLayout(main_layout)
        vbox.addWidget(self.log_output)
        main_widget.setLayout(vbox)

        # 헤더/그리드 스타일
        self.excel_view.setStyleSheet("""
        QHeaderView::section {
            background-color: #3F4A73;
            color: white;
            font-weight: bold;
            border: 1px solid #555;
        }
        QTableWidget {
            gridline-color: #AAAAAA;
        }
        """)

    def log(self, msg):
        """로그 메시지를 출력합니다."""
        if hasattr(self, 'log_output'):
            self.log_output.append(msg)
        print(msg)

    def open_excel(self):
        """엑셀 파일을 엽니다."""
        path, _ = QFileDialog.getOpenFileName(
            self, "엑셀 파일 선택", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        self.log(f"[작업] 엑셀 파일 열기: {path}")
        try:
            self.excel_path = path
            self.json_path, _ = (
                self.excel_service.load_excel(path, self.excel_view)
            )
            self.log(f"[작업] JSON 파일 저장: {self.json_path}")
        except Exception as e:
            self.log(f"[오류] 엑셀 파일 열기 실패: {e}")
            QMessageBox.critical(
                self,
                "엑셀 파일 오류",
                f"엑셀 파일을 불러올 수 없습니다:\n{e}"
            )

    def on_cell_changed(self, row, col):
        """셀이 변경되었을 때 호출됩니다."""
        if not self.json_path:
            return
        data_dict = self._widget_to_json_schema()
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump(data_dict, f, ensure_ascii=False, indent=2)
        self.log(f"[작업] 셀 변경: ({row},{col}) → JSON 동기화")

    def _widget_to_json_schema(self):
        """테이블 위젯의 내용을 JSON 스키마로 변환합니다."""
        table = self.excel_view
        result = {
            "meta": {},
            "items": [],
            "discounts": [],
            "summary": {},
            "comments": ""
        }
        current_category = None
        for row in range(table.rowCount()):
            a_item = table.item(row, 0)
            d_item = table.item(row, 3)
            a = a_item.text() if a_item else ""
            d = d_item.text() if d_item else ""
            # 1) 섹션 헤더
            if a and not d:
                current_category = a.strip()
                continue
            # 2) 품목 행
            if a and d:
                try:
                    item = {
                        "category": current_category,
                        "description": a.strip(),
                        "unit_price": (
                            float(table.item(row, 1).text())
                            if table.item(row, 1)
                            and table.item(row, 1).text()
                            else 0
                        ),
                        "quantity": (
                            float(table.item(row, 2).text())
                            if table.item(row, 2)
                            and table.item(row, 2).text()
                            else 0
                        ),
                        "unit_count": float(table.item(row, 3).text()),
                        "amount": (
                            float(table.item(row, 4).text())
                            if table.columnCount() > 4
                            and table.item(row, 4)
                            and table.item(row, 4).text()
                            else None
                        )
                    }
                    result["items"].append(item)
                except Exception:
                    continue
                continue
            # 3) 요약/할인/합계 행 (A열 비어있고 D열에 숫자)
            if not a and d:
                try:
                    val = float(d)
                    if val < 0:
                        result["discounts"].append({
                            "description": "",
                            "amount": -val
                        })
                    else:
                        if "subtotal" not in result["summary"]:
                            result["summary"]["subtotal"] = val
                        elif "tax_amount" not in result["summary"]:
                            result["summary"]["tax_amount"] = val
                        else:
                            result["summary"]["total_due"] = val
                except Exception:
                    continue
                continue
        return result

    def ask_gpt(self):
        """GPT에 질문을 보냅니다."""
        user_q = self.chat_input.toPlainText().strip()
        if not user_q or not self.json_path:
            return
        with open(self.json_path, "r", encoding="utf-8") as f:
            quotation_json = f.read()
        messages = [
            {
                "role": "system",
                "content": "아래 견적서 JSON을 참고해 질문에 답변해 주세요."
            },
            {
                "role": "user",
                "content": (
                    f"견적서 데이터:\n```json\n{quotation_json}\n```\n"
                    f"질문: {user_q}"
                )
            }
        ]
        answer = ask_gpt_api(messages, GPT_API_KEY, GPT_MODEL)
        self.chat_output.append(f"<b>질문:</b> {user_q}")
        self.chat_output.append("<b>GPT:</b>")
        self.chat_output.append(answer)
        self.chat_input.clear()
        self.log("[작업] GPT 질문 전송 및 응답 수신 완료")

    def eventFilter(self, obj, event):
        """이벤트 필터: 엔터키로 질문 전송"""
        if (
            obj == self.chat_input and
            event.type() == event.KeyPress
        ):
            if (
                event.key() == Qt.Key_Return and
                event.modifiers() == Qt.NoModifier
            ):
                self.ask_gpt()
                return True
        return super().eventFilter(obj, event) 