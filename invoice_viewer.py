import sys
import os
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QLabel,
    QLineEdit, QFormLayout, QTextEdit
)
from PyQt5.QtCore import Qt, QSizeF
from PyQt5.QtGui import QColor, QBrush
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QTextDocument
from reportlab_invoice_sample import save_invoice_to_pdf

class InvoiceTemplateViewer(QMainWindow):
    def __init__(self, json_dir):
        super().__init__()
        self.json_dir = json_dir
        self.setWindowTitle("견적서 템플릿 뷰어")
        self.setGeometry(200, 200, 1200, 900)
        self.init_ui()

    def init_ui(self):
        central = QWidget()
        vbox = QVBoxLayout()

        # 상단 고정 정보
        self.header_label = QLabel(
            "<b>U-STUDIO</b><br>"
            "363, Gangnam-daero<br>"
            "Seocho-gu, Seoul, Republic of Korea<br>"
            "Phone: +82-2-549-2048 / +82-10-9870-1024<br>"
            "Fax: +82-2-539-2047<br>"
            "VAT Number: 451-81-00624<br>"
            "Banking Number: WOORI BANK 1005-903-051608<br>"
            "SWIFT CODE: HVBKKRSEXXX"
        )
        self.header_label.setAlignment(Qt.AlignLeft)
        vbox.addWidget(self.header_label)

        # 기본 정보 입력 폼
        form_layout = QFormLayout()
        self.quotation_no = QLineEdit()
        self.quotation_date = QLineEdit()
        self.company_name = QLineEdit()
        self.payment_date = QLineEdit()
        self.ship_to = QLineEdit()
        form_layout.addRow("견적번호:", self.quotation_no)
        form_layout.addRow("견적일자:", self.quotation_date)
        form_layout.addRow("거래처명:", self.company_name)
        form_layout.addRow("Payment date:", self.payment_date)
        form_layout.addRow("Ship To:", self.ship_to)
        vbox.addLayout(form_layout)

        # 파일 선택
        hbox = QHBoxLayout()
        self.file_label = QLabel("파일을 선택하세요")
        self.btn_open = QPushButton("JSON 파일 열기")
        self.btn_open.clicked.connect(self.open_json_file)
        hbox.addWidget(self.file_label)
        hbox.addWidget(self.btn_open)
        vbox.addLayout(hbox)

        # 표 영역
        self.table = QTableWidget()
        vbox.addWidget(self.table)

        # 행 추가/삭제 버튼
        btn_layout = QHBoxLayout()
        self.btn_add_row = QPushButton("행 추가")
        self.btn_delete_row = QPushButton("행 삭제")
        self.btn_add_row.clicked.connect(self.add_row)
        self.btn_delete_row.clicked.connect(self.delete_row)
        btn_layout.addWidget(self.btn_add_row)
        btn_layout.addWidget(self.btn_delete_row)
        vbox.addLayout(btn_layout)

        # 합계 정보
        total_layout = QHBoxLayout()
        self.total_amount = QLineEdit()
        self.tax_amount = QLineEdit()
        self.grand_total = QLineEdit()
        
        total_layout.addWidget(QLabel("합계금액:"))
        total_layout.addWidget(self.total_amount)
        total_layout.addWidget(QLabel("세액:"))
        total_layout.addWidget(self.tax_amount)
        total_layout.addWidget(QLabel("총액:"))
        total_layout.addWidget(self.grand_total)
        vbox.addLayout(total_layout)

        # 하단 고정 정보
        self.footer_label = QLabel(
            "<b>OTHER COMMENTS</b><br>"
            "If you have any questions about this quotation, please contact<br>"
            "U-STUDIO BS, support@ustudio.co.kr<br>"
            "<b>Thank You For Your Business!</b>"
        )
        self.footer_label.setAlignment(Qt.AlignLeft)
        vbox.addWidget(self.footer_label)

        # OTHER COMMENTS 입력란 추가
        self.other_comments_edit = QTextEdit()
        self.other_comments_edit.setPlaceholderText("OTHER COMMENTS를 입력하세요...")
        vbox.addWidget(self.other_comments_edit)

        central.setLayout(vbox)
        self.setCentralWidget(central)

        # 저장 버튼 추가
        self.add_save_button()
        self.add_save_pdf_button()

    def open_json_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "JSON 파일 선택", self.json_dir, "JSON Files (*.json)"
        )
        if file_path:
            self.file_label.setText(os.path.basename(file_path))
            self.last_loaded_filename = os.path.splitext(os.path.basename(file_path))[0]  # 확장자 제외
            self.load_json(file_path)

    def load_json(self, file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.loaded_json_data = data  # 원본 JSON 데이터 저장
        # 기본 정보 설정
        self.quotation_no.setText(data.get("견적번호", ""))
        self.quotation_date.setText(data.get("견적일자", ""))
        self.company_name.setText(data.get("거래처명", ""))
        self.payment_date.setText(data.get("payment_date", ""))
        self.ship_to.setText(data.get("ship_to", ""))

        # 카테고리/세부항목 구조 지원
        categories = data.get("카테고리")
        headers = ["품목명", "수량", "단가", "금액", "비고"]
        if categories:
            rows = []
            row_types = []
            for cat in categories:
                rows.append({"품목명": cat.get("category", ""), "수량": "", "단가": "", "금액": cat.get("amount", ""), "비고": ""})
                row_types.append("cat")
                for item in cat.get("items", []):
                    rows.append({
                        "품목명": item.get("품목명", ""),
                        "수량": item.get("수량", ""),
                        "단가": item.get("단가", ""),
                        "금액": item.get("금액", ""),
                        "비고": item.get("비고", "")
                    })
                    row_types.append("item")
            self.table.setColumnCount(len(headers))
            self.table.setRowCount(len(rows))
            self.table.setHorizontalHeaderLabels(headers)
            sky_blue = QBrush(QColor(230, 243, 255))
            white = QBrush(QColor(255, 255, 255))
            for row, (item, rtype) in enumerate(zip(rows, row_types)):
                for col, key in enumerate(headers):
                    val = str(item.get(key, ""))
                    cell = QTableWidgetItem(val)
                    if rtype == "cat":
                        cell.setBackground(sky_blue)
                    else:
                        cell.setBackground(white)
                    self.table.setItem(row, col, cell)
        else:
            # 기존 품목 배열 구조 지원
            self.table.setColumnCount(len(headers))
            self.table.setRowCount(len(data.get("품목", [])))
            self.table.setHorizontalHeaderLabels(headers)
            for row, item in enumerate(data.get("품목", [])):
                for col, key in enumerate(headers):
                    val = str(item.get(key, ""))
                    cell = QTableWidgetItem(val)
                    cell.setBackground(QBrush(QColor(255, 255, 255)))
                    self.table.setItem(row, col, cell)

        # 합계 정보 설정
        self.total_amount.setText(str(data.get("합계금액", "")))
        self.tax_amount.setText(str(data.get("세액", "")))
        self.grand_total.setText(str(data.get("총액", "")))

        # OTHER COMMENTS 입력란 값 세팅
        self.other_comments_edit.setPlainText(data.get("other_comments", ""))

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

    def add_row(self):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for c in range(self.table.columnCount()):
            self.table.setItem(row, c, QTableWidgetItem(""))

    def delete_row(self):
        selected = self.table.currentRow()
        if selected >= 0:
            self.table.removeRow(selected)

    def save_table_to_json(self, save_path):
        data = {
            "견적번호": self.quotation_no.text(),
            "견적일자": self.quotation_date.text(),
            "거래처명": self.company_name.text(),
            "payment_date": self.payment_date.text(),
            "ship_to": self.ship_to.text(),
            "품목": []
        }

        for row in range(self.table.rowCount()):
            item = {
                "품목명": self.table.item(row, 0).text() if self.table.item(row, 0) else "",
                "수량": self.table.item(row, 1).text() if self.table.item(row, 1) else "",
                "단가": self.table.item(row, 2).text() if self.table.item(row, 2) else "",
                "금액": self.table.item(row, 3).text() if self.table.item(row, 3) else "",
                "비고": self.table.item(row, 4).text() if self.table.item(row, 4) else ""
            }
            if any(item.values()):  # 빈 행은 저장하지 않음
                data["품목"].append(item)

        data["합계금액"] = self.total_amount.text()
        data["세액"] = self.tax_amount.text()
        data["총액"] = self.grand_total.text()

        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def add_save_button(self):
        btn_save = QPushButton("JSON 저장")
        btn_save.clicked.connect(self.save_json_dialog)
        self.centralWidget().layout().addWidget(btn_save)

    def save_json_dialog(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "JSON 파일로 저장", self.json_dir, "JSON Files (*.json)"
        )
        if file_path:
            self.save_table_to_json(file_path)

    def add_save_pdf_button(self):
        btn_save_pdf = QPushButton("PDF로 저장")
        btn_save_pdf.clicked.connect(self.save_pdf_dialog)
        self.centralWidget().layout().addWidget(btn_save_pdf)

    def save_pdf_dialog(self):
        default_name = getattr(self, 'last_loaded_filename', '견적서') + '.pdf'
        file_path, _ = QFileDialog.getSaveFileName(
            self, "PDF 파일로 저장", os.path.join(self.json_dir, default_name), "PDF Files (*.pdf)"
        )
        if file_path:
            data = {
                "견적번호": self.quotation_no.text(),
                "견적일자": self.quotation_date.text(),
                "거래처명": self.company_name.text(),
                "payment_date": self.payment_date.text(),
                "ship_to": self.ship_to.text(),
                "품목": [],
            }
            if hasattr(self, 'loaded_json_data') and "카테고리" in self.loaded_json_data:
                data["카테고리"] = self.loaded_json_data["카테고리"]
            else:
                for row in range(self.table.rowCount()):
                    item = {
                        "품목명": self.table.item(row, 0).text() if self.table.item(row, 0) else "",
                        "수량": self.table.item(row, 1).text() if self.table.item(row, 1) else "",
                        "단가": self.table.item(row, 2).text() if self.table.item(row, 2) else "",
                        "금액": self.table.item(row, 3).text() if self.table.item(row, 3) else "",
                        "비고": self.table.item(row, 4).text() if self.table.item(row, 4) else ""
                    }
                    if any(item.values()):
                        data["품목"].append(item)
            # OTHER COMMENTS 입력란 값 반영
            data['other_comments'] = self.other_comments_edit.toPlainText()
            save_invoice_to_pdf(data, file_path)

    def generate_invoice_html(self):
        date = self.quotation_date.text()
        quote_no = self.quotation_no.text()
        payment_date = self.payment_date.text()
        ship_to = self.ship_to.text()
        total = self.total_amount.text()
        tax = self.tax_amount.text()
        grand = self.grand_total.text()
        try:
            total_val = float(total.replace(',', ''))
            tax_val = float(tax.replace(',', ''))
            tax_rate = f"{tax_val / total_val * 100:.3f}%" if total_val else ""
        except Exception:
            tax_rate = ""
        html = """
        <style>
          body { font-size:12pt; font-family:'Malgun Gothic', Arial, sans-serif; }
          .invoice-container { margin-left:40px; margin-right:40px; }
          table { border-collapse:collapse; width:100%; }
          th, td { padding:6px; font-family:'Malgun Gothic', Arial, sans-serif; }
          .items th { background-color:#e6f3ff; border:1px solid #999; }
          .items td { border:1px solid #999; }
          .header-table td { vertical-align:top; }
          .meta-table td { padding:4px 8px; }
          .footer-left th { background-color:#4f81bd; color:white; }
          .footer-right td { border:1px solid #999; padding:4px 8px; }
          .yellow { background-color:#ffff00; }
        </style>
        <div class=\"invoice-container\">"""
        html += f"""
        <table class=\"header-table\">\n  <tr>\n    <td width=\"60%\">\n      <h2>U-STUDIO</h2>\n      363, Gangnam-daero<br>\n      Seocho-gu, Seoul, Republic of Korea<br>\n      Phone: +82-2-549-2048 / +82-10-9870-1024<br>\n      Fax: +82-2-539-2047<br>\n      VAT Number: 451-81-00624<br>\n      Banking Number: WOORI BANK 1005-903-051608<br>\n      SWIFT CODE: HVBKKRSEXXX\n    </td>\n    <td width=\"40%\" align=\"right\">\n      <h1>INVOICE</h1>\n      <table class=\"meta-table\">\n        <tr><td><b>Date:</b></td><td>{date}</td></tr>\n        <tr><td><b>Quotation #:</b></td><td>{quote_no}</td></tr>\n        <tr><td><b>Payment date:</b></td><td>{payment_date}</td></tr>\n        <tr><td><b>Ship To:</b></td><td>{ship_to}</td></tr>\n      </table>\n    </td>\n  </tr>\n</table>\n"""
        html += """
        <table class=\"items\" style=\"width:100%;\">
          <colgroup>
            <col style='width:40%;'>
            <col style='width:12%;'>
            <col style='width:14%;'>
            <col style='width:18%;'>
            <col style='width:16%;'>
          </colgroup>
          <tr class=\"items\">
            <th style='border:1px solid #999;'>품목명</th>
            <th style='border:1px solid #999;'>수량</th>
            <th style='border:1px solid #999;'>단가</th>
            <th style='border:1px solid #999;'>금액</th>
            <th style='border:1px solid #999;'>비고</th>
          </tr>
        """
        for row in range(self.table.rowCount()):
            is_cat = (
                (self.table.item(row, 1) and self.table.item(row, 1).text() == "") and
                (self.table.item(row, 2) and self.table.item(row, 2).text() == "")
            )
            bg = "#e6f3ff" if is_cat else "#ffffff"
            html += f"<tr style='background-color:{bg};'>"
            for col in range(self.table.columnCount()):
                val = self.table.item(row, col).text() if self.table.item(row, col) else ""
                if col == 0:
                    html += f"<td style='width:40%;border:1px solid #999;'>{val}</td>"
                elif col == 1:
                    html += f"<td style='width:12%;border:1px solid #999;'>{val}</td>"
                elif col == 2:
                    html += f"<td style='width:14%;border:1px solid #999;'>{val}</td>"
                elif col == 3:
                    html += f"<td style='width:18%;border:1px solid #999;'>{val}</td>"
                else:
                    html += f"<td style='width:16%;border:1px solid #999;'>{val}</td>"
            html += "</tr>"
        html += "</table>"
        html += f"""
        <br><br>
        <table width=\"100%\">
          <tr>
            <td width=\"60%\" valign=\"top\">
              <table class=\"footer-left\" width=\"100%\">
                <tr><th>OTHER COMMENTS</th></tr>
                <tr><td>&nbsp;</td></tr>
                <tr><td>If you have any questions about this quotation, please contact<br>
                  U-STUDIO BS, support@ustudio.co.kr
                </td></tr>
                <tr><td><i>Thank You For Your Business!</i></td></tr>
              </table>
            </td>
            <td width=\"40%\" valign=\"top\">
              <table class=\"footer-right\" width=\"100%\">
                <tr><td><b>Total</b></td><td align=\"right\" class=\"yellow\">{total}</td></tr>
                <tr><td><b>Tax rate</b></td><td align=\"right\">{tax_rate}</td></tr>
                <tr><td><b>Tax due</b></td><td align=\"right\" class=\"yellow\">{tax}</td></tr>
                <tr><td><b>Other</b></td><td align=\"right\">–</td></tr>
                <tr><td><b>TOTAL Due</b></td><td align=\"right\" class=\"yellow\">{grand}</td></tr>
              </table>
            </td>
          </tr>
        </table>
        """
        html += "</div>"
        return html

if __name__ == "__main__":
    app = QApplication(sys.argv)
    json_dir = "./PDFtoJSON"
    viewer = InvoiceTemplateViewer(json_dir)
    viewer.show()
    sys.exit(app.exec_()) 