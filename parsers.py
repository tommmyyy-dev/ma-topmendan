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
