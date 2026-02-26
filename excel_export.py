"""Excel出力: 分析結果をExcelファイルに書き出す"""

from __future__ import annotations

import io
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from analyzer import AnalysisResult


# スタイル定義
_HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
_HEADER_FONT = Font(name="Yu Gothic", bold=True, color="FFFFFF", size=10)
_PRIORITY_FILLS = {
    "A": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    "B": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    "C": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
}
_RISK_FILLS = {
    "high": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    "medium": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    "low": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
}
_BODY_FONT = Font(name="Yu Gothic", size=10)
_WRAP_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
_THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _apply_style(ws, row, col, font=None, fill=None):
    cell = ws.cell(row=row, column=col)
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    cell.alignment = _WRAP_ALIGNMENT
    cell.border = _THIN_BORDER


def export_to_excel(result: AnalysisResult) -> bytes:
    """分析結果をExcelバイト列として返す"""
    wb = Workbook()

    # ---- Sheet 1: 質問リスト ----
    ws1 = wb.active
    ws1.title = "質問リスト"

    headers1 = [
        ("No.", 6),
        ("カテゴリ", 22),
        ("優先度", 8),
        ("質問事項", 55),
        ("質問の意図・目的", 35),
        ("背景・根拠", 35),
        ("関連資料", 22),
        ("フォローアップ", 35),
        ("回答メモ", 30),
    ]

    for col_idx, (header, width) in enumerate(headers1, 1):
        cell = ws1.cell(row=1, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _WRAP_ALIGNMENT
        cell.border = _THIN_BORDER
        ws1.column_dimensions[get_column_letter(col_idx)].width = width

    # 優先度順にソート (A > B > C)
    priority_order = {"A": 0, "B": 1, "C": 2}
    sorted_questions = sorted(
        result.questions,
        key=lambda q: (q.category_code, priority_order.get(q.priority, 3)),
    )

    for i, q in enumerate(sorted_questions, 1):
        row = i + 1
        ws1.cell(row=row, column=1, value=i)
        ws1.cell(row=row, column=2, value=q.category_name)
        ws1.cell(row=row, column=3, value=q.priority)
        ws1.cell(row=row, column=4, value=q.question)
        ws1.cell(row=row, column=5, value=q.intent)
        ws1.cell(row=row, column=6, value=q.background)
        ws1.cell(row=row, column=7, value="\n".join(q.source_documents))
        ws1.cell(row=row, column=8, value="\n".join(q.follow_up_points))
        ws1.cell(row=row, column=9, value="")  # 回答メモ（空欄）

        for col_idx in range(1, 10):
            _apply_style(ws1, row, col_idx, font=_BODY_FONT)

        # 優先度セルに色付け
        fill = _PRIORITY_FILLS.get(q.priority)
        if fill:
            ws1.cell(row=row, column=3).fill = fill

    ws1.freeze_panes = "A2"
    ws1.auto_filter.ref = f"A1:I{len(sorted_questions) + 1}"

    # ---- Sheet 2: 主要論点・リスク ----
    ws2 = wb.create_sheet("主要論点・リスク")

    headers2 = [
        ("No.", 6),
        ("論点", 30),
        ("詳細", 55),
        ("リスクレベル", 12),
        ("関連カテゴリ", 20),
    ]

    for col_idx, (header, width) in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _WRAP_ALIGNMENT
        cell.border = _THIN_BORDER
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    for i, ki in enumerate(result.key_issues, 1):
        row = i + 1
        risk_label = {"high": "高", "medium": "中", "low": "低"}.get(ki.risk_level, ki.risk_level)
        cat_names = ", ".join(ki.related_categories)

        ws2.cell(row=row, column=1, value=i)
        ws2.cell(row=row, column=2, value=ki.title)
        ws2.cell(row=row, column=3, value=ki.description)
        ws2.cell(row=row, column=4, value=risk_label)
        ws2.cell(row=row, column=5, value=cat_names)

        for col_idx in range(1, 6):
            _apply_style(ws2, row, col_idx, font=_BODY_FONT)

        fill = _RISK_FILLS.get(ki.risk_level)
        if fill:
            ws2.cell(row=row, column=4).fill = fill

    ws2.freeze_panes = "A2"

    # ---- Sheet 3: 企業サマリー ----
    ws3 = wb.create_sheet("企業サマリー")
    ws3.column_dimensions["A"].width = 15
    ws3.column_dimensions["B"].width = 80

    company_url = getattr(result, "company_url", "") or ""
    summary_data = [
        ("対象企業", result.company_name),
        ("企業URL", company_url),
        ("作成日", datetime.now().strftime("%Y年%m月%d日")),
        ("質問数", f"{len(result.questions)}問"),
        ("論点数", f"{len(result.key_issues)}件"),
        ("", ""),
        ("企業サマリー", result.summary),
    ]

    for i, (label, value) in enumerate(summary_data, 1):
        cell_a = ws3.cell(row=i, column=1, value=label)
        cell_a.font = Font(name="Yu Gothic", bold=True, size=10)
        cell_a.alignment = _WRAP_ALIGNMENT
        cell_b = ws3.cell(row=i, column=2, value=value)
        cell_b.font = _BODY_FONT
        cell_b.alignment = _WRAP_ALIGNMENT

    # バイト列で返す
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
