"""M&A Agent Team 分析ツール — Streamlit UI"""

from __future__ import annotations

import hashlib
import json
import os
from datetime import datetime
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

from agents import DEFAULT_BUYER_PROFILE, run_full_analysis
from parsers import SUPPORTED_EXTENSIONS, parse_file, extract_financial_summary

load_dotenv()

# ---------------------------------------------------------------------------
# ページ設定
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="M&A Agent Team 分析",
    page_icon="🤖",
    layout="wide",
    menu_items={
        "Get help": None,
        "Report a Bug": None,
        "About": "M&A Agent Team 分析ツール v2.0\n\n"
        "IM・財務資料をアップロードすると、3つのAIエージェントが"
        "並列で案件分析・市場調査・Red Teamレビューを実行します。",
    },
)

# ---------------------------------------------------------------------------
# パスワード認証
# ---------------------------------------------------------------------------
def _check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True

    correct_password = ""
    try:
        correct_password = st.secrets["passwords"]["app_password"]
    except (KeyError, FileNotFoundError):
        correct_password = os.environ.get("APP_PASSWORD", "")

    if not correct_password:
        return True

    expected_token = hashlib.sha256(
        f"ma-topmendan:{correct_password}".encode()
    ).hexdigest()[:16]

    if st.query_params.get("t") == expected_token:
        st.session_state.authenticated = True
        return True

    st.markdown("## 🔐 ログイン")
    password = st.text_input("パスワード", type="password", key="login_password")
    remember = st.checkbox("ログイン情報を保存する")

    if st.button("ログイン", type="primary"):
        if password == correct_password:
            st.session_state.authenticated = True
            if remember:
                st.query_params["t"] = expected_token
            st.rerun()
        else:
            st.error("パスワードが正しくありません。")
    return False


if not _check_password():
    st.stop()

# ---------------------------------------------------------------------------
# カスタムCSS
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
    .main-header { font-size: 1.6rem; font-weight: 700; margin-bottom: 0.2rem; }
    .sub-header { font-size: 0.95rem; color: #666; margin-bottom: 1.5rem; }
    [data-testid="stSidebar"] { min-width: 280px; }
    [data-testid="stSidebar"] .stButton > button {
        padding-top: 0.15rem; padding-bottom: 0.15rem;
        min-height: 0; font-size: 0.78rem; line-height: 1.2;
    }
    .verdict-go { background: #d4edda; color: #155724; padding: 12px 20px;
                   border-radius: 8px; font-size: 1.3rem; font-weight: 700; }
    .verdict-cgo { background: #fff3cd; color: #856404; padding: 12px 20px;
                   border-radius: 8px; font-size: 1.3rem; font-weight: 700; }
    .verdict-nogo { background: #f8d7da; color: #721c24; padding: 12px 20px;
                    border-radius: 8px; font-size: 1.3rem; font-weight: 700; }
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# データ永続化 (Google Sheets / ローカルJSON)
# ---------------------------------------------------------------------------
DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)
THREADS_FILE = DATA_DIR / "threads.json"

SPREADSHEET_ID = "1Ee5cSvpHTa9qXPAT_uWGBSJ8Gcm_2Wxyj1t9lRfImBA"


def _get_gspread_client():
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        creds_info = dict(st.secrets.get("gcp_service_account", {}))
        if not creds_info:
            return None
        creds = Credentials.from_service_account_info(
            creds_info,
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        return gspread.authorize(creds)
    except Exception:
        return None


def _load_threads_from_gsheet() -> dict:
    gc = _get_gspread_client()
    if gc is None:
        return {}
    try:
        ws = gc.open_by_key(SPREADSHEET_ID).worksheet("threads")
        rows = ws.get_all_values()
        threads = {}
        for row in rows[1:]:
            if len(row) >= 3 and row[0]:
                try:
                    threads[row[0]] = json.loads(row[2])
                except (json.JSONDecodeError, IndexError):
                    pass
        return threads
    except Exception:
        return {}


def _save_thread_to_gsheet(thread_id: str, thread_data: dict) -> bool:
    gc = _get_gspread_client()
    if gc is None:
        return False
    try:
        ws = gc.open_by_key(SPREADSHEET_ID).worksheet("threads")
        rows = ws.get_all_values()
        company = thread_data.get("company_name", "不明")
        json_str = json.dumps(thread_data, ensure_ascii=False)
        for i, row in enumerate(rows[1:], start=2):
            if row[0] == thread_id:
                ws.update(f"A{i}:C{i}", [[thread_id, company, json_str]])
                return True
        ws.append_row([thread_id, company, json_str], value_input_option="RAW")
        return True
    except Exception:
        return False


def _delete_thread_from_gsheet(thread_id: str) -> bool:
    gc = _get_gspread_client()
    if gc is None:
        return False
    try:
        ws = gc.open_by_key(SPREADSHEET_ID).worksheet("threads")
        rows = ws.get_all_values()
        for i, row in enumerate(rows[1:], start=2):
            if row[0] == thread_id:
                ws.delete_rows(i)
                return True
        return False
    except Exception:
        return False


def _load_threads() -> dict:
    threads = _load_threads_from_gsheet()
    if threads:
        return threads
    if THREADS_FILE.exists():
        try:
            return json.loads(THREADS_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {}


def _save_threads(threads: dict):
    THREADS_FILE.write_text(
        json.dumps(threads, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ---------------------------------------------------------------------------
# セッション初期化
# ---------------------------------------------------------------------------
if "threads" not in st.session_state:
    st.session_state.threads = _load_threads()
if "current_thread" not in st.session_state:
    st.session_state.current_thread = None

# ---------------------------------------------------------------------------
# サイドバー
# ---------------------------------------------------------------------------
PROVIDERS = {
    "Anthropic (Claude)": {
        "models": ["claude-sonnet-4-20250514", "claude-haiku-4-5-20251001"],
    },
    "OpenAI (GPT)": {
        "models": ["gpt-4o", "gpt-4o-mini"],
    },
}

with st.sidebar:
    if st.button("＋ 新しい案件", use_container_width=True, type="primary"):
        st.session_state.current_thread = None
        st.rerun()

    st.markdown("---")

    st.markdown("### API設定")
    provider = st.selectbox("AIプロバイダー", list(PROVIDERS.keys()), index=0)
    api_key = st.text_input(f"{provider} API Key", type="password")
    model = st.selectbox("モデル", PROVIDERS[provider]["models"], index=0)

    with st.expander("買い手プロファイル"):
        buyer_profile_input = st.text_area(
            "編集可能",
            value=DEFAULT_BUYER_PROFILE,
            height=200,
            label_visibility="collapsed",
        )

    st.markdown("---")

    st.markdown("### 案件一覧")
    threads = st.session_state.threads
    if threads:
        for tid in sorted(
            threads.keys(),
            key=lambda k: threads[k].get("created_at", ""),
            reverse=True,
        ):
            t = threads[tid]
            label = t.get("company_name", "不明な案件")
            created = t.get("created_at", "")[:10]
            is_active = st.session_state.current_thread == tid

            col_sel, col_del = st.columns([5, 1])
            with col_sel:
                if st.button(
                    f"{'▶ ' if is_active else ''}{label}\n{created}",
                    key=f"t_{tid}",
                    use_container_width=True,
                    disabled=is_active,
                ):
                    st.session_state.current_thread = tid
                    st.rerun()
            with col_del:
                if st.button("🗑", key=f"d_{tid}", help="削除"):
                    del st.session_state.threads[tid]
                    _save_threads(st.session_state.threads)
                    _delete_thread_from_gsheet(tid)
                    if st.session_state.current_thread == tid:
                        st.session_state.current_thread = None
                    st.rerun()
    else:
        st.caption("まだ案件がありません")


# ---------------------------------------------------------------------------
# 結果表示
# ---------------------------------------------------------------------------
def _extract_verdict(red_team_text: str) -> str:
    upper = red_team_text[:500].upper()
    if "NO-GO" in upper:
        return "NO-GO"
    if "CONDITIONAL" in upper:
        return "CONDITIONAL-GO"
    if "GO" in upper:
        return "GO"
    return ""


def _show_result(data: dict):
    company = data.get("company_name", "不明")
    st.markdown(f"## 🤖 {company}")

    # Red Team 判定バッジ
    verdict = _extract_verdict(data.get("red_team", ""))
    if verdict:
        css = {"GO": "verdict-go", "CONDITIONAL-GO": "verdict-cgo", "NO-GO": "verdict-nogo"}
        st.markdown(
            f'<div class="{css.get(verdict, "")}">&nbsp;Red Team判定: {verdict}</div>',
            unsafe_allow_html=True,
        )
        st.markdown("")

    # トークン
    in_tok = data.get("total_input_tokens", 0)
    out_tok = data.get("total_output_tokens", 0)
    if in_tok or out_tok:
        st.caption(f"トークン: 入力 {in_tok:,} / 出力 {out_tok:,}")

    # Markdownダウンロード
    full_md = f"# {company} — M&A Agent Team 分析レポート\n\n"
    full_md += f"生成日時: {data.get('created_at', '')}\n\n"
    full_md += "---\n\n# 📋 案件分析\n\n" + data.get("main_analysis", "")
    full_md += "\n\n---\n\n# 🌐 市場調査\n\n" + data.get("market_research", "")
    full_md += "\n\n---\n\n# 🔴 Red Team レビュー\n\n" + data.get("red_team", "")

    st.download_button(
        "📥 レポートをダウンロード (.md)",
        data=full_md,
        file_name=f"{company}_Agent分析_{data.get('created_at', '')[:10]}.md",
        mime="text/markdown",
    )

    # 3タブで表示
    tab1, tab2, tab3 = st.tabs([
        "📋 案件分析（IM・財務・DD）",
        "🌐 市場調査",
        "🔴 Red Team",
    ])
    with tab1:
        st.markdown(data.get("main_analysis", "*分析結果がありません*"))
    with tab2:
        st.markdown(data.get("market_research", "*市場調査結果がありません*"))
    with tab3:
        st.markdown(data.get("red_team", "*Red Teamレビューがありません*"))


# ---------------------------------------------------------------------------
# メインエリア
# ---------------------------------------------------------------------------

# 既存スレッド表示
if (
    st.session_state.current_thread
    and st.session_state.current_thread in st.session_state.threads
):
    data = st.session_state.threads[st.session_state.current_thread]

    # v1（旧バージョン）の場合
    if data.get("version") != 2:
        st.info(
            f"**{data.get('company_name', '不明')}** — 旧バージョンの分析結果です。"
            "「新しい案件」から再分析してください。"
        )
    else:
        _show_result(data)

# 新規分析画面
else:
    st.markdown(
        '<p class="main-header">M&A Agent Team 分析</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="sub-header">IM・財務資料をアップロードすると、3つのAIエージェントが'
        "並列で分析します</p>",
        unsafe_allow_html=True,
    )

    _upload_types = [ext.lstrip(".") for ext in SUPPORTED_EXTENSIONS]
    uploaded_files = st.file_uploader(
        "資料をアップロード（最大10ファイル）",
        type=_upload_types,
        accept_multiple_files=True,
        help="PDF, Excel, Word, CSV, PowerPoint等に対応",
    )

    if uploaded_files:
        if len(uploaded_files) > 10:
            st.error("最大10ファイルまでです。")
            st.stop()

        st.markdown(f"**{len(uploaded_files)}件** のファイルが選択されています")

        if st.button("🚀 Agent Team 分析開始", type="primary"):
            if not api_key:
                st.error(f"サイドバーで {provider} の API Key を設定してください。")
                st.stop()

            # ファイル解析
            with st.status("資料を解析中...", expanded=True) as status:
                parsed_docs = []
                raw_files = []
                for f in uploaded_files:
                    st.write(f"📄 {f.name}")
                    file_bytes = f.read()
                    raw_files.append((f.name, file_bytes))
                    doc = parse_file(f.name, file_bytes)
                    parsed_docs.append(doc)
                    for w in doc.warnings:
                        st.warning(f"  ⚠ {f.name}: {w}")

                valid_docs = [d for d in parsed_docs if d.text.strip()]
                if not valid_docs:
                    st.error("テキストを抽出できるファイルがありませんでした。")
                    st.stop()

                fin_summary = extract_financial_summary(raw_files)
                status.update(
                    label=f"✅ {len(valid_docs)}件の資料を解析完了",
                    state="complete",
                )

            # Agent Team 分析
            with st.status("🤖 Agent Team 分析中...", expanded=True) as status:
                try:
                    result = run_full_analysis(
                        api_key=api_key,
                        provider=provider,
                        model=model,
                        documents=valid_docs,
                        financial_summary=fin_summary,
                        buyer_profile=buyer_profile_input,
                        progress_cb=lambda msg: st.write(msg),
                    )
                except Exception as e:
                    st.error(f"分析エラー: {e}")
                    st.stop()

                status.update(label="✅ 分析完了!", state="complete")

            # スレッド保存
            thread_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            thread_data = {
                "version": 2,
                "company_name": result.company_name,
                "main_analysis": result.main_analysis,
                "market_research": result.market_research,
                "red_team": result.red_team,
                "total_input_tokens": result.total_input_tokens,
                "total_output_tokens": result.total_output_tokens,
                "created_at": datetime.now().isoformat(),
            }
            st.session_state.threads[thread_id] = thread_data
            _save_threads(st.session_state.threads)
            _save_thread_to_gsheet(thread_id, thread_data)

            st.session_state.current_thread = thread_id
            st.rerun()

    else:
        st.markdown("---")
        st.markdown(
            """
### 📎 資料をアップロードして始めましょう

**3つのAIエージェントが並列で分析します:**

| エージェント | 分析内容 |
|---|---|
| 📋 **案件分析** | IM分析、財務分析（感応度・バリュエーション・のれん償却）、DD計画、トップ面談質問 |
| 🌐 **市場調査** | Web検索による市場規模、M&A動向、競合分析 |
| 🔴 **Red Team** | 批判的レビュー、Deal Breaker、ダウンサイドシナリオ、GO/NO-GO判定 |

**対応ファイル**: PDF, Excel, Word, CSV, PowerPoint, HTML, JSON, テキスト等
"""
        )
