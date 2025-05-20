import os
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_LEFT, TA_RIGHT

# 한글 폰트 등록 (윈도우 기준)
font_dir = os.path.join(os.environ.get('WINDIR', 'C:/Windows'), 'Fonts')
pdfmetrics.registerFont(TTFont('MalgunGothic', os.path.join(font_dir, 'malgun.ttf')))
pdfmetrics.registerFont(TTFont('MalgunGothic-Bold', os.path.join(font_dir, 'malgunbd.ttf')))

def save_invoice_to_pdf(data, save_path):
    doc = SimpleDocTemplate(
        save_path,
        pagesize=A4,
        leftMargin=30,
        rightMargin=30,
        topMargin=30,
        bottomMargin=30
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='BlueHeader', fontName='MalgunGothic', fontSize=10,
        textColor=colors.white, backColor=colors.HexColor('#4f81bd'),
        alignment=1, spaceAfter=0, spaceBefore=0, leading=12
    ))
    styles.add(ParagraphStyle(
        name='NormalKor', fontName='MalgunGothic', fontSize=10, leading=14
    ))

    story = []

    # 상단 정보
    header_data = [
        [
            Paragraph(
                "<b>U-STUDIO</b><br/>363, Gangnam-daero<br/>Seocho-gu, Seoul, Republic of Korea<br/>"
                "Phone: +82-2-549-2048 / +82-10-9870-1024<br/>Fax: +82-2-539-2047<br/>"
                "VAT Number: 451-81-00624<br/>Banking Number: WOORI BANK 1005-903-051608<br/>"
                "SWIFT CODE: HVBKKRSEXXX",
                styles['NormalKor']
            ),
            "",
            Paragraph(
                f"<b>INVOICE</b><br/><br/>DATE: {data.get('견적일자', '')}" \
                f"<br/>QUOTATION #: {data.get('견적번호', '')}" \
                f"<br/>Payment date: {data.get('payment_date', '')}" \
                f"<br/>SHIP TO: {data.get('ship_to', '')}",
                styles['NormalKor']
            )
        ]
    ]
    header_table = Table(header_data, colWidths=[250, 20, 220])
    header_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (0, 0)),
        ('SPAN', (2, 0), (2, 0)),
        ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 12))

    # 표 데이터 준비
    table_header = [
        [
            Paragraph("DESCRIPTION", styles['BlueHeader']),
            Paragraph("UNIT KRW", styles['BlueHeader']),
            Paragraph("QTY", styles['BlueHeader']),
            Paragraph("AMOUNT", styles['BlueHeader']),
            Paragraph("REMARKS", styles['BlueHeader'])
        ]
    ]
    table_data = []
    row_styles = []
    if "카테고리" in data:
        for cat in data["카테고리"]:
            # 카테고리 행 (하늘색)
            table_data.append([
                Paragraph(
                    str(cat.get("category", "")),
                    ParagraphStyle('cat', fontName='MalgunGothic', fontSize=10, leading=12, wordWrap='CJK')
                ),
                "", "", cat.get("amount", ""), ""
            ])
            row_styles.append("cat")
            for item in cat.get("items", []):
                table_data.append([
                    Paragraph(
                        str(item.get("품목명", "")),
                        ParagraphStyle('desc', fontName='MalgunGothic', fontSize=10, leading=12, wordWrap='CJK')
                    ),
                    item.get("단가", ""),
                    item.get("수량", ""),
                    item.get("금액", ""),
                    item.get("비고", "")
                ])
                row_styles.append("item")
    else:
        for item in data.get("품목", []):
            table_data.append([
                Paragraph(
                    str(item.get("품목명", "")),
                    ParagraphStyle('desc', fontName='MalgunGothic', fontSize=10, leading=12, wordWrap='CJK')
                ),
                item.get("단가", ""),
                item.get("수량", ""),
                item.get("금액", ""),
                item.get("비고", "")
            ])
            row_styles.append("item")
    while len(table_data) < 20:
        table_data.append(["", "", "", "", ""])
        row_styles.append("item")
    full_table_data = table_header + table_data
    table_style = TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'MalgunGothic'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4f81bd')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])
    # 카테고리/항목별 배경색 적용
    for i, rtype in enumerate(row_styles):
        if rtype == "cat":
            table_style.add('BACKGROUND', (0, i + 1), (-1, i + 1), colors.HexColor('#e6f3ff'))
        else:
            table_style.add('BACKGROUND', (0, i + 1), (-1, i + 1), colors.white)

    # 세부내역 표 rowHeights=20 제거 (자동 높이)
    table = Table(full_table_data, colWidths=[200, 70, 50, 100, 100])
    table.setStyle(table_style)
    story.append(table)
    story.append(Spacer(1, 12))

    # OTHER COMMENTS 박스 (1열 2행)
    other_comments_text = data.get('other_comments', '')
    other_comments_inner_table = Table(
        [
            [Paragraph("OTHER COMMENTS", styles['BlueHeader'])],
            [Paragraph(other_comments_text, styles['NormalKor'])]
        ],
        colWidths=[220],
        style=TableStyle([
            ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#4f81bd')),
            ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
            ('FONTNAME', (0, 0), (0, 0), 'MalgunGothic'),
            ('FONTSIZE', (0, 0), (0, 0), 10),
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('LINEBELOW', (0, 0), (0, 0), 1, colors.black),
        ])
    )

    # 합계 박스 (summary_table) 생성 코드 그대로 유지
    def format_won(val):
        try:
            v = int(str(val).replace(",", ""))
            return f"￦ {v:,.0f}"
        except Exception:
            return f"￦ {val}" if val not in ("", "-") else val
    total = data.get("합계금액", "")
    tax = data.get("세액", "")
    grand = data.get("총액", "")
    summary_table_data = [
        [Paragraph("TOTAL", ParagraphStyle('sumlabel', fontName='MalgunGothic', fontSize=10, alignment=TA_LEFT)),
         Paragraph(format_won(total), ParagraphStyle('sumval', fontName='MalgunGothic', fontSize=10, alignment=TA_RIGHT))],
        [Paragraph("Tax rate", ParagraphStyle('sumlabel', fontName='MalgunGothic', fontSize=10, alignment=TA_LEFT)),
         Paragraph("10.000%", ParagraphStyle('sumval', fontName='MalgunGothic', fontSize=10, alignment=TA_RIGHT))],
        [Paragraph("Tax due", ParagraphStyle('sumlabel', fontName='MalgunGothic', fontSize=10, alignment=TA_LEFT)),
         Paragraph(format_won(tax), ParagraphStyle('sumval', fontName='MalgunGothic', fontSize=10, alignment=TA_RIGHT))],
        [Paragraph("Other", ParagraphStyle('sumlabel', fontName='MalgunGothic', fontSize=10, alignment=TA_LEFT)),
         Paragraph("-", ParagraphStyle('sumval', fontName='MalgunGothic', fontSize=10, alignment=TA_RIGHT))],
        [Paragraph("TOTAL Due", ParagraphStyle('sumlabel', fontName='MalgunGothic-Bold', fontSize=11, alignment=TA_LEFT)),
         Paragraph(format_won(grand), ParagraphStyle('sumval', fontName='MalgunGothic-Bold', fontSize=11, alignment=TA_RIGHT))]
    ]
    summary_table = Table(
        summary_table_data,
        colWidths=[90, 110],
        style=TableStyle([
            ('FONTNAME', (0, 0), (-1, -2), 'MalgunGothic'),
            ('FONTNAME', (0, -1), (-1, -1), 'MalgunGothic-Bold'),
            ('FONTSIZE', (0, 0), (-1, -2), 10),
            ('FONTSIZE', (0, -1), (-1, -1), 11),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, -1), (1, -1), colors.yellow),
            ('TEXTCOLOR', (0, -1), (1, -1), colors.black),
        ]),
        hAlign='RIGHT'
    )

    # 하단 2열 Table로 좌: OTHER COMMENTS, 우: 합계 박스
    bottom_row_table = Table(
        [[other_comments_inner_table, summary_table]],
        colWidths=[320, 200],  # 세부내역표 colWidths 합과 동일하게
        style=TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ])
    )
    story.append(Spacer(1, 18))
    story.append(bottom_row_table)
    story.append(Spacer(1, 8))
    # 안내 문구
    story.append(Paragraph(
        "If you have any questions about this quotation, please contact<br/>"
        "U-STUDIO BS, support@ustudio.co.kr",
        styles['NormalKor']
    ))
    story.append(Spacer(1, 6))
    story.append(Paragraph("Thank You For Your Business!", styles['NormalKor']))

    doc.build(story) 