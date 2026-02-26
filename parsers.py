"""ドキュメントパーサー: 各種ファイル形式からテキストを抽出する"""

from __future__ import annotations

import csv
import io
import json as json_lib
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------------------
# 対応拡張子一覧（app.py側のfile_uploaderと合わせること）
# ---------------------------------------------------------------------------
SUPPORTED_EXTENSIONS = [
    ".pdf",
    ".xlsx", ".xls", ".ods",       # スプレッドシート系
    ".csv", ".tsv",                 # テキスト表形式
    ".docx", ".doc", ".rtf",       # ワープロ系
    ".txt", ".text", ".md",        # プレーンテキスト / Markdown
    ".json",                        # JSON
    ".html", ".htm",               # HTML
    ".pptx",                        # PowerPoint
]


@dataclass
class ParsedDocument:
    filename: str
    doc_type: str
    text: str
    page_count: int = 0
    sheet_names: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def parse_file(file_name: str, file_bytes: bytes) -> ParsedDocument:
    """ファイル拡張子に応じて適切なパーサーを呼び出す"""
    ext = Path(file_name).suffix.lower()

    dispatch = {
        ".pdf": _parse_pdf,
        ".xlsx": _parse_excel,
        ".xls": _parse_excel,
        ".ods": _parse_ods,
        ".csv": _parse_csv,
        ".tsv": _parse_tsv,
        ".docx": _parse_word,
        ".doc": _parse_doc_legacy,
        ".rtf": _parse_rtf,
        ".txt": _parse_text,
        ".text": _parse_text,
        ".md": _parse_text,
        ".json": _parse_json,
        ".html": _parse_html,
        ".htm": _parse_html,
        ".pptx": _parse_pptx,
    }

    parser = dispatch.get(ext)
    if parser:
        return parser(file_name, file_bytes)

    return ParsedDocument(
        filename=file_name,
        doc_type="Unknown",
        text="",
        warnings=[f"未対応の形式です: {ext}"],
    )


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
def _parse_pdf(name: str, data: bytes) -> ParsedDocument:
    import pdfplumber

    warnings: list[str] = []
    pages_text: list[str] = []

    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                tables = page.extract_tables()
                table_text = ""
                for table in tables:
                    for row in table:
                        cells = [str(c) if c else "" for c in row]
                        table_text += " | ".join(cells) + "\n"
                combined = text
                if table_text:
                    combined += "\n[表データ]\n" + table_text
                pages_text.append(combined)

            if not any(t.strip() for t in pages_text):
                warnings.append("テキスト抽出ができませんでした（スキャン画像の可能性）")

            return ParsedDocument(
                filename=name,
                doc_type="PDF",
                text="\n\n---\n\n".join(pages_text),
                page_count=len(pdf.pages),
                warnings=warnings,
            )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="PDF", text="",
            warnings=[f"PDF解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# Excel (.xlsx / .xls)
# ---------------------------------------------------------------------------
def _parse_excel(name: str, data: bytes) -> ParsedDocument:
    from openpyxl import load_workbook

    warnings: list[str] = []
    sheets_text: list[str] = []
    sheet_names: list[str] = []

    try:
        wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
        for ws_name in wb.sheetnames:
            sheet_names.append(ws_name)
            ws = wb[ws_name]
            rows: list[str] = []
            row_count = 0
            for row in ws.iter_rows(values_only=True):
                row_count += 1
                if row_count > 5000:
                    warnings.append(f"シート '{ws_name}' は5000行で切り捨てました")
                    break
                cells = [str(c) if c is not None else "" for c in row]
                if any(cells):
                    rows.append(" | ".join(cells))
            sheets_text.append(f"【シート: {ws_name}】\n" + "\n".join(rows))
        wb.close()

        return ParsedDocument(
            filename=name,
            doc_type="Excel",
            text="\n\n".join(sheets_text),
            sheet_names=sheet_names,
            warnings=warnings,
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="Excel", text="",
            warnings=[f"Excel解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# ODS (LibreOffice / Google Sheets export)
# ---------------------------------------------------------------------------
def _parse_ods(name: str, data: bytes) -> ParsedDocument:
    """ODS形式をpandasで読み込む。odfpy が必要。"""
    try:
        import pandas as pd

        sheets = pd.read_excel(io.BytesIO(data), engine="odf", sheet_name=None)
        parts: list[str] = []
        sheet_names: list[str] = []
        for sname, df in sheets.items():
            sheet_names.append(str(sname))
            text = df.to_csv(sep="|", index=False)
            parts.append(f"【シート: {sname}】\n{text}")

        return ParsedDocument(
            filename=name,
            doc_type="ODS",
            text="\n\n".join(parts),
            sheet_names=sheet_names,
        )
    except ImportError:
        return ParsedDocument(
            filename=name, doc_type="ODS", text="",
            warnings=["ODS読み込みには odfpy パッケージが必要です (pip install odfpy)"],
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="ODS", text="",
            warnings=[f"ODS解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------
def _parse_csv(name: str, data: bytes) -> ParsedDocument:
    try:
        text = _decode_bytes(data)
        reader = csv.reader(io.StringIO(text))
        rows = [" | ".join(row) for row in reader if any(row)]
        return ParsedDocument(
            filename=name, doc_type="CSV",
            text="\n".join(rows[:5000]),
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="CSV", text="",
            warnings=[f"CSV解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# TSV
# ---------------------------------------------------------------------------
def _parse_tsv(name: str, data: bytes) -> ParsedDocument:
    try:
        text = _decode_bytes(data)
        reader = csv.reader(io.StringIO(text), delimiter="\t")
        rows = [" | ".join(row) for row in reader if any(row)]
        return ParsedDocument(
            filename=name, doc_type="TSV",
            text="\n".join(rows[:5000]),
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="TSV", text="",
            warnings=[f"TSV解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# Word (.docx)
# ---------------------------------------------------------------------------
def _parse_word(name: str, data: bytes) -> ParsedDocument:
    from docx import Document

    warnings: list[str] = []
    try:
        doc = Document(io.BytesIO(data))
        paragraphs: list[str] = []
        for para in doc.paragraphs:
            if para.text.strip():
                prefix = ""
                if para.style and para.style.name and para.style.name.startswith("Heading"):
                    prefix = "## "
                paragraphs.append(prefix + para.text)

        for table in doc.tables:
            trows: list[str] = []
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                trows.append(" | ".join(cells))
            if trows:
                paragraphs.append("[表データ]\n" + "\n".join(trows))

        return ParsedDocument(
            filename=name,
            doc_type="Word",
            text="\n".join(paragraphs),
            page_count=len(doc.paragraphs) // 30 or 1,
            warnings=warnings,
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="Word", text="",
            warnings=[f"Word解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# Word (.doc) - レガシー形式
# ---------------------------------------------------------------------------
def _parse_doc_legacy(name: str, data: bytes) -> ParsedDocument:
    """古い.doc形式。テキスト部分のバイナリ抽出を試みる。"""
    warnings: list[str] = []
    try:
        # バイナリからテキスト部分を抽出（簡易的）
        text = data.decode("utf-8", errors="ignore")
        # 制御文字を除去
        cleaned = "".join(c for c in text if c.isprintable() or c in "\n\r\t")
        lines = [l.strip() for l in cleaned.splitlines() if l.strip() and len(l.strip()) > 3]
        if not lines:
            warnings.append(".doc形式のテキスト抽出は限定的です。.docx形式への変換を推奨します")
        return ParsedDocument(
            filename=name,
            doc_type="Word (legacy)",
            text="\n".join(lines[:3000]),
            warnings=warnings,
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="Word (legacy)", text="",
            warnings=[f".doc解析エラー: {e}。.docx形式への変換を推奨します"],
        )


# ---------------------------------------------------------------------------
# RTF
# ---------------------------------------------------------------------------
def _parse_rtf(name: str, data: bytes) -> ParsedDocument:
    """RTF形式。striprtf パッケージがあれば使用、なければ簡易抽出。"""
    try:
        from striprtf.striprtf import rtf_to_text
        text = rtf_to_text(data.decode("utf-8", errors="ignore"))
        return ParsedDocument(filename=name, doc_type="RTF", text=text)
    except ImportError:
        # striprtf がない場合は簡易抽出
        raw = data.decode("utf-8", errors="ignore")
        import re
        text = re.sub(r"[\\{}\[\]]", " ", raw)
        text = re.sub(r"\\[a-z]+\d*\s?", "", text)
        cleaned = " ".join(text.split())
        return ParsedDocument(
            filename=name, doc_type="RTF", text=cleaned,
            warnings=["RTFの高精度解析には striprtf パッケージが必要です (pip install striprtf)"],
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="RTF", text="",
            warnings=[f"RTF解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# プレーンテキスト / Markdown
# ---------------------------------------------------------------------------
def _parse_text(name: str, data: bytes) -> ParsedDocument:
    ext = Path(name).suffix.lower()
    doc_type = "Markdown" if ext == ".md" else "テキスト"
    try:
        text = _decode_bytes(data)
        return ParsedDocument(filename=name, doc_type=doc_type, text=text[:100000])
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type=doc_type, text="",
            warnings=[f"テキスト解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# JSON
# ---------------------------------------------------------------------------
def _parse_json(name: str, data: bytes) -> ParsedDocument:
    try:
        text = _decode_bytes(data)
        obj = json_lib.loads(text)
        pretty = json_lib.dumps(obj, ensure_ascii=False, indent=2)
        return ParsedDocument(
            filename=name, doc_type="JSON",
            text=pretty[:100000],
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="JSON", text="",
            warnings=[f"JSON解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------
def _parse_html(name: str, data: bytes) -> ParsedDocument:
    """HTMLからテキストを抽出。BeautifulSoup があれば使用。"""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(data, "html.parser")
        # スクリプトとスタイルを除去
        for tag in soup(["script", "style"]):
            tag.decompose()
        text = soup.get_text(separator="\n", strip=True)
        return ParsedDocument(filename=name, doc_type="HTML", text=text[:100000])
    except ImportError:
        # bs4 がない場合は簡易抽出
        import re
        raw = _decode_bytes(data)
        text = re.sub(r"<script[^>]*>.*?</script>", "", raw, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r"<[^>]+>", " ", text)
        text = " ".join(text.split())
        return ParsedDocument(
            filename=name, doc_type="HTML", text=text[:100000],
            warnings=["HTML高精度解析には beautifulsoup4 が必要です (pip install beautifulsoup4)"],
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="HTML", text="",
            warnings=[f"HTML解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# PowerPoint (.pptx)
# ---------------------------------------------------------------------------
def _parse_pptx(name: str, data: bytes) -> ParsedDocument:
    """PowerPointからテキストを抽出。python-pptx が必要。"""
    try:
        from pptx import Presentation

        prs = Presentation(io.BytesIO(data))
        slides_text: list[str] = []
        for i, slide in enumerate(prs.slides, 1):
            texts: list[str] = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        t = para.text.strip()
                        if t:
                            texts.append(t)
                if shape.has_table:
                    for row in shape.table.rows:
                        cells = [cell.text.strip() for cell in row.cells]
                        texts.append(" | ".join(cells))
            if texts:
                slides_text.append(f"【スライド {i}】\n" + "\n".join(texts))

        return ParsedDocument(
            filename=name,
            doc_type="PowerPoint",
            text="\n\n".join(slides_text),
            page_count=len(prs.slides),
        )
    except ImportError:
        return ParsedDocument(
            filename=name, doc_type="PowerPoint", text="",
            warnings=["PowerPoint読み込みには python-pptx が必要です (pip install python-pptx)"],
        )
    except Exception as e:
        return ParsedDocument(
            filename=name, doc_type="PowerPoint", text="",
            warnings=[f"PowerPoint解析エラー: {e}"],
        )


# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------
def _decode_bytes(data: bytes) -> str:
    """バイト列を文字列にデコード。日本語エンコーディングを自動検出。"""
    for enc in ("utf-8", "utf-8-sig", "shift_jis", "cp932", "euc-jp", "iso-2022-jp"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


# ---------------------------------------------------------------------------
# 財務CSV自動抽出
# ---------------------------------------------------------------------------
def _parse_num(s: str) -> int | None:
    """数値文字列をintに変換"""
    s = s.strip().replace(",", "")
    if not s or s == "-":
        return None
    try:
        return int(s)
    except ValueError:
        try:
            return int(float(s))
        except ValueError:
            return None


def _readable_yen(val: int) -> str:
    """金額を読みやすい日本語表記で返す（例: 123,456,789（約1.2億円））"""
    abs_val = abs(val)
    if abs_val >= 100_000_000:
        readable = f"約{val / 100_000_000:.1f}億円"
    elif abs_val >= 10_000_000:
        readable = f"約{val / 10_000:.0f}万円"
    elif abs_val >= 10_000:
        readable = f"約{val / 10_000:.0f}万円"
    else:
        return f"{val:,}円"
    return f"{val:,}円（{readable}）"


def _find_row(
    rows: list[list[str]], patterns: list[str], cols: tuple[int, ...] = (0,)
) -> list[str] | None:
    """パターンに一致し数値データを含む行を返す"""
    for row in rows:
        if not row or len(row) < 4:
            continue
        for c in cols:
            if c >= len(row):
                continue
            label = row[c].strip()
            for p in patterns:
                if label == p:
                    for i in range(3, min(len(row), 20)):
                        if _parse_num(row[i]) is not None:
                            return row
    return None


_PL_ITEMS: list[tuple[str, list[str]]] = [
    ("売上高", ["売上高合計", "売上高計"]),
    ("売上原価", ["売上原価合計", "売上原価計"]),
    ("売上総利益", ["売上総利益", "売上総利益合計"]),
    ("販管費", ["販売費及び一般管理費合計", "販売費及一般管理費合計", "販管費合計"]),
    ("営業利益", ["営業利益", "営業利益合計"]),
    ("経常利益", ["経常利益", "経常利益合計"]),
    ("税引前利益", ["税引前当期純利益", "税引前当期純利益合計"]),
    ("当期純利益", ["当期純利益", "当期純利益合計"]),
]

_BS_ITEMS: list[tuple[str, list[str], tuple[int, ...]]] = [
    ("総資産", ["資産の部合計"], (0,)),
    ("流動資産", ["流動資産合計"], (0,)),
    ("現預金", ["現金及び預金合計", "現金及び預金", "現金及預金合計"], (0, 1)),
    ("固定資産", ["固定資産合計"], (0,)),
    ("負債合計", ["負債の部合計"], (0,)),
    ("流動負債", ["流動負債合計"], (0,)),
    ("固定負債", ["固定負債合計"], (0,)),
    ("有利子負債", ["長期借入金合計", "長期借入金"], (0, 1)),
    ("純資産", ["純資産の部合計"], (0,)),
]


def _format_pl(fname: str, text: str) -> str:
    """損益計算書CSVから主要PLデータを抽出してフォーマット"""
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if len(rows) < 3:
        return ""

    header = rows[0]
    month_cols: list[tuple[int, str]] = []
    total_col: int | None = None
    for i in range(3, len(header)):
        h = header[i].strip()
        if h == "合計":
            total_col = i
        elif h and h != "決算整理":
            month_cols.append((i, h))

    lines = [f"■ 損益計算書（月次推移）- {fname}", "単位: 円", ""]

    # 年間合計（読みやすい単位付き）
    lines.append("【年間合計】")
    found_any = False
    for display, patterns in _PL_ITEMS:
        row = _find_row(rows, patterns, cols=(0,))
        if not row:
            continue
        found_any = True
        val = None
        if total_col is not None and total_col < len(row):
            val = _parse_num(row[total_col])
        if val is not None:
            lines.append(f"  {display}: {_readable_yen(val)}")

    if not found_any:
        return ""

    # 月次推移（全PL項目）
    lines.append("")
    lines.append("【月次推移】")
    # ヘッダー行
    month_labels = [mlabel for _, mlabel in month_cols]
    lines.append("科目 | " + " | ".join(month_labels))
    for item_display, item_patterns in _PL_ITEMS:
        row = _find_row(rows, item_patterns, cols=(0,))
        if not row:
            continue
        vals = []
        for ci, _ in month_cols:
            v = _parse_num(row[ci]) if ci < len(row) else None
            vals.append(f"{v:,}" if v is not None else "-")
        lines.append(f"{item_display} | " + " | ".join(vals))

    return "\n".join(lines)


def _format_bs(fname: str, text: str) -> str:
    """貸借対照表CSVから主要BSデータを抽出してフォーマット"""
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if len(rows) < 3:
        return ""

    header = rows[0]
    month_cols: list[tuple[int, str]] = []
    for i in range(3, len(header)):
        h = header[i].strip()
        if h:
            month_cols.append((i, h))

    if not month_cols:
        return ""

    latest_col, latest_label = month_cols[-1]
    lines = [f"■ 貸借対照表 - {fname}", "単位: 円", f"基準日: {latest_label}", ""]

    lines.append("【残高】")
    found_any = False
    for display, patterns, cols in _BS_ITEMS:
        row = _find_row(rows, patterns, cols=cols)
        if row and latest_col < len(row):
            val = _parse_num(row[latest_col])
            if val is not None:
                lines.append(f"  {display}: {_readable_yen(val)}")
                found_any = True

    if not found_any:
        return ""

    # 月次推移（総資産・純資産・現預金）
    for item_display, item_patterns, item_cols in [
        ("総資産", ["資産の部合計"], (0,)),
        ("純資産", ["純資産の部合計"], (0,)),
        ("現預金", ["現金及び預金合計", "現金及び預金", "現金及預金合計"], (0, 1)),
    ]:
        row = _find_row(rows, item_patterns, cols=item_cols)
        if not row:
            continue
        parts = []
        for ci, mlabel in month_cols:
            v = _parse_num(row[ci]) if ci < len(row) else None
            parts.append(f"{mlabel}={v:,}" if v is not None else f"{mlabel}=-")
        lines.append("")
        lines.append(f"月次{item_display}: {' / '.join(parts)}")

    return "\n".join(lines)


def extract_financial_summary(files: list[tuple[str, bytes]]) -> str:
    """財務CSVから主要データを自動抽出し、LLM用の構造化テキストで返す"""
    parts: list[str] = []
    for fname, data in files:
        if not fname.lower().endswith(".csv"):
            continue
        try:
            text = _decode_bytes(data)
        except Exception:
            continue
        if "損益" in fname:
            s = _format_pl(fname, text)
            if s:
                parts.append(s)
        elif "貸借" in fname:
            s = _format_bs(fname, text)
            if s:
                parts.append(s)

    if not parts:
        return ""
    return (
        "【財務データ（CSVから自動抽出・正確な数値）】\n"
        "以下は財務CSVからプログラムで正確に抽出した数値です。\n"
        "financial_analysisセクションにはこの数値をそのまま使用してください。\n"
        "IMの記載と異なる場合でもCSVデータを優先してください。\n\n"
        + "\n\n".join(parts)
    )
