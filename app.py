"""M&A トップ面談 質問リスト自動生成ツール - Streamlit UI"""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

from analyzer import (
    PROVIDERS,
    AnalysisResult,
    FinancialAnalysis,
    analyze_documents,
)
from excel_export import export_to_excel
from parsers import SUPPORTED_EXTENSIONS, ParsedDocument, parse_file

load_dotenv()

# ---------------------------------------------------------------------------
# ページ設定
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="M&A トップ面談 質問ジェネレーター",
    page_icon="📋",
    layout="wide",
    menu_items={
        "Get help": None,
        "Report a Bug": None,
        "About": "M&A トップ面談 質問リスト自動生成ツール v1.2\n\n"
                 "IM・財務諸表・試算表などの資料をアップロードし、"
                 "対象企業のトップ面談用の質問リスト・論点を自動生成します。",
    },
)

# ---------------------------------------------------------------------------
# パスワード認証
# ---------------------------------------------------------------------------
def _check_password() -> bool:
    """ログイン画面を表示し、認証済みならTrueを返す。"""
    if st.session_state.get("authenticated"):
        return True

    # パスワードは Streamlit secrets または環境変数から取得
    # secrets.toml: [passwords] app_password = "your_password"
    correct_password = ""
    try:
        correct_password = st.secrets["passwords"]["app_password"]
    except (KeyError, FileNotFoundError):
        correct_password = os.environ.get("APP_PASSWORD", "")

    # パスワード未設定の場合は認証スキップ（ローカル開発用）
    if not correct_password:
        return True

    st.markdown("## 🔐 ログイン")
    st.markdown("このアプリを利用するにはパスワードを入力してください。")
    password = st.text_input("パスワード", type="password", key="login_password")

    if st.button("ログイン", type="primary"):
        if password == correct_password:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません。")
    return False


if not _check_password():
    st.stop()

# ---------------------------------------------------------------------------
# カスタムCSS（ハンバーガーメニュー日本語化含む）
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
    .main-header {
        font-size: 1.6rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 0.95rem;
        color: #666;
        margin-bottom: 1.5rem;
    }
    .company-url {
        font-size: 0.9rem;
        margin-top: -0.5rem;
        margin-bottom: 1rem;
    }
    .company-url a {
        color: #1a73e8;
        text-decoration: none;
    }
    .company-url a:hover { text-decoration: underline; }

    /* ハンバーガーメニュー内の英語テキストを日本語に置換 */
    [data-testid="stMainMenu"] ul li:nth-child(1) span { font-size: 0; }
    [data-testid="stMainMenu"] ul li:nth-child(1) span::after {
        content: "再実行"; font-size: 14px;
    }
    [data-testid="stMainMenu"] ul li:nth-child(2) span { font-size: 0; }
    [data-testid="stMainMenu"] ul li:nth-child(2) span::after {
        content: "設定"; font-size: 14px;
    }

    /* スレッドボタンのスタイル */
    .thread-btn {
        display: block;
        width: 100%;
        padding: 8px 12px;
        margin-bottom: 4px;
        border: none;
        border-radius: 6px;
        text-align: left;
        cursor: pointer;
        font-size: 0.9rem;
    }
    .thread-active {
        background: #e3f2fd;
        font-weight: 600;
    }

    /* サイドバー幅 */
    [data-testid="stSidebar"] { min-width: 280px; }
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# データ永続化 (JSONファイル)
# ---------------------------------------------------------------------------
DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)
THREADS_FILE = DATA_DIR / "threads.json"


def _load_threads() -> dict:
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


def _financial_to_dict(fa: FinancialAnalysis) -> dict:
    return {
        "pl_trends": [
            {
                "period": p.period, "revenue": p.revenue,
                "operating_profit": p.operating_profit,
                "ordinary_profit": p.ordinary_profit,
                "net_income": p.net_income, "ebitda": p.ebitda, "unit": p.unit,
            }
            for p in fa.pl_trends
        ],
        "monthly_quarterly_trends": [
            {
                "period": m.period, "revenue": m.revenue,
                "operating_profit": m.operating_profit,
                "unit": m.unit, "trend_type": m.trend_type,
            }
            for m in fa.monthly_quarterly_trends
        ],
        "bs_trends": [
            {
                "period": b.period, "total_assets": b.total_assets,
                "total_liabilities": b.total_liabilities,
                "net_assets": b.net_assets,
                "cash_and_deposits": b.cash_and_deposits,
                "interest_bearing_debt": b.interest_bearing_debt, "unit": b.unit,
            }
            for b in fa.bs_trends
        ],
        "key_metrics": [
            {"metric_name": km.metric_name, "values": km.values}
            for km in fa.key_metrics
        ],
        "financial_comments": fa.financial_comments,
    }


def _result_to_dict(r: AnalysisResult) -> dict:
    return {
        "company_name": r.company_name,
        "company_url": r.company_url,
        "summary": r.summary,
        "questions": [
            {
                "category_code": q.category_code,
                "category_name": q.category_name,
                "question": q.question,
                "intent": q.intent,
                "background": q.background,
                "source_documents": q.source_documents,
                "priority": q.priority,
                "follow_up_points": q.follow_up_points,
            }
            for q in r.questions
        ],
        "key_issues": [
            {
                "title": ki.title,
                "description": ki.description,
                "risk_level": ki.risk_level,
                "related_categories": ki.related_categories,
            }
            for ki in r.key_issues
        ],
        "financial_analysis": _financial_to_dict(r.financial_analysis),
        "input_tokens": r.input_tokens,
        "output_tokens": r.output_tokens,
        "created_at": datetime.now().isoformat(),
    }


def _dict_to_financial(d: dict) -> FinancialAnalysis:
    from analyzer import BSTrend, KeyMetric, MonthlyQuarterlyTrend, PLTrend

    fa = d.get("financial_analysis", {})
    if not fa:
        return FinancialAnalysis()
    return FinancialAnalysis(
        pl_trends=[
            PLTrend(
                period=p.get("period", ""), revenue=p.get("revenue"),
                operating_profit=p.get("operating_profit"),
                ordinary_profit=p.get("ordinary_profit"),
                net_income=p.get("net_income"), ebitda=p.get("ebitda"),
                unit=p.get("unit", "千円"),
            )
            for p in fa.get("pl_trends", [])
        ],
        monthly_quarterly_trends=[
            MonthlyQuarterlyTrend(
                period=m.get("period", ""), revenue=m.get("revenue"),
                operating_profit=m.get("operating_profit"),
                unit=m.get("unit", "千円"),
                trend_type=m.get("trend_type", "monthly"),
            )
            for m in fa.get("monthly_quarterly_trends", [])
        ],
        bs_trends=[
            BSTrend(
                period=b.get("period", ""), total_assets=b.get("total_assets"),
                total_liabilities=b.get("total_liabilities"),
                net_assets=b.get("net_assets"),
                cash_and_deposits=b.get("cash_and_deposits"),
                interest_bearing_debt=b.get("interest_bearing_debt"),
                unit=b.get("unit", "千円"),
            )
            for b in fa.get("bs_trends", [])
        ],
        key_metrics=[
            KeyMetric(metric_name=km.get("metric_name", ""), values=km.get("values", []))
            for km in fa.get("key_metrics", [])
        ],
        financial_comments=fa.get("financial_comments", ""),
    )


def _dict_to_result(d: dict) -> AnalysisResult:
    from analyzer import KeyIssue, Question

    return AnalysisResult(
        company_name=d.get("company_name", ""),
        company_url=d.get("company_url", ""),
        summary=d.get("summary", ""),
        questions=[
            Question(
                category_code=q["category_code"],
                category_name=q["category_name"],
                question=q["question"],
                intent=q["intent"],
                background=q["background"],
                source_documents=q["source_documents"],
                priority=q["priority"],
                follow_up_points=q["follow_up_points"],
            )
            for q in d.get("questions", [])
        ],
        key_issues=[
            KeyIssue(
                title=ki["title"],
                description=ki["description"],
                risk_level=ki["risk_level"],
                related_categories=ki["related_categories"],
            )
            for ki in d.get("key_issues", [])
        ],
        financial_analysis=_dict_to_financial(d),
        input_tokens=d.get("input_tokens", 0),
        output_tokens=d.get("output_tokens", 0),
    )


# ---------------------------------------------------------------------------
# セッション初期化
# ---------------------------------------------------------------------------
if "threads" not in st.session_state:
    st.session_state.threads = _load_threads()  # {thread_id: {...}}
if "current_thread" not in st.session_state:
    st.session_state.current_thread = None  # thread_id or None (= 新規)
if "result" not in st.session_state:
    st.session_state.result = None

# ---------------------------------------------------------------------------
# サイドバー
# ---------------------------------------------------------------------------
with st.sidebar:
    # ---- 新規作成ボタン ----
    if st.button("＋ 新しい案件の分析を作成", use_container_width=True, type="primary"):
        st.session_state.current_thread = None
        st.session_state.result = None
        st.rerun()

    st.markdown("---")

    # ---- 案件スレッド一覧 ----
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
            q_count = len(t.get("questions", []))
            is_active = st.session_state.current_thread == tid

            btn_label = f"{'▶ ' if is_active else ''}{label}　({q_count}問)"
            if st.button(
                btn_label,
                key=f"thread_{tid}",
                use_container_width=True,
                disabled=is_active,
            ):
                st.session_state.current_thread = tid
                st.session_state.result = _dict_to_result(t)
                st.rerun()

            if is_active:
                # 削除ボタン
                if st.button("🗑 この案件を削除", key=f"del_{tid}", type="secondary"):
                    del st.session_state.threads[tid]
                    _save_threads(st.session_state.threads)
                    st.session_state.current_thread = None
                    st.session_state.result = None
                    st.rerun()
    else:
        st.caption("まだ案件がありません")

    st.markdown("---")

    # ---- API設定 ----
    st.markdown("### API設定")

    provider = st.selectbox(
        "AIプロバイダー",
        list(PROVIDERS.keys()),
        index=1,
    )
    provider_info = PROVIDERS[provider]

    # APIキーは Secrets > 環境変数 > 手入力 の優先順で取得
    _secrets_key_map = {
        "ANTHROPIC_API_KEY": "anthropic_api_key",
        "OPENAI_API_KEY": "openai_api_key",
    }
    _default_key = ""
    try:
        _default_key = st.secrets.get(
            _secrets_key_map.get(provider_info["env_key"], ""), ""
        )
    except (KeyError, FileNotFoundError):
        pass
    if not _default_key:
        _default_key = os.environ.get(provider_info["env_key"], "")

    api_key = st.text_input(
        f"{provider} API Key",
        value=_default_key,
        type="password",
        help="Secrets に保存済みなら自動入力されます",
    )

    model = st.selectbox("モデル", provider_info["models"], index=0)

    st.markdown("---")
    st.markdown("### 使い方")
    st.markdown(
        """
1. 「新しい案件の分析を作成」をクリック
2. 資料ファイルをアップロード（最大10件）
3. 「分析開始」を押す
4. 結果を確認 → Excelダウンロード

**対応形式**: PDF, Excel, Word, CSV, TSV, PowerPoint, HTML, JSON, テキスト, Markdown, ODS, RTF
"""
    )

# ---------------------------------------------------------------------------
# メインエリア
# ---------------------------------------------------------------------------
result: AnalysisResult | None = st.session_state.result

# ---- 新規分析画面（スレッド未選択 or 新規作成） ----
if st.session_state.current_thread is None and result is None:
    st.markdown(
        '<p class="main-header">M&A トップ面談 質問ジェネレーター</p>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="sub-header">IM・財務諸表・試算表などの資料をアップロードすると、'
        "トップ面談用の質問リスト・論点を自動生成します</p>",
        unsafe_allow_html=True,
    )

    # 対応拡張子から先頭のドットを除いてリスト化
    _upload_types = [ext.lstrip(".") for ext in SUPPORTED_EXTENSIONS]

    uploaded_files = st.file_uploader(
        "資料をアップロード（最大10ファイル）",
        type=_upload_types,
        accept_multiple_files=True,
        help="PDF, Excel, Word, CSV, PowerPoint, HTML, JSON, テキスト等に対応。最大10ファイル。",
    )

    if uploaded_files:
        if len(uploaded_files) > 10:
            st.error("アップロードは1度に最大10ファイルまでです。ファイルを減らしてください。")
            st.stop()

        st.markdown(f"**{len(uploaded_files)}件のファイルが選択されています**")

        if st.button("🔍 分析開始", type="primary"):
            if not api_key:
                st.error(f"サイドバーで {provider} の API Key を設定してください。")
                st.stop()

            # ファイル解析
            with st.status("資料を解析中...", expanded=True) as status:
                parsed_docs: list[ParsedDocument] = []
                for f in uploaded_files:
                    st.write(f"📄 {f.name} を解析中...")
                    doc = parse_file(f.name, f.read())
                    parsed_docs.append(doc)
                    if doc.warnings:
                        for w in doc.warnings:
                            st.warning(f"  ⚠ {f.name}: {w}")

                valid_docs = [d for d in parsed_docs if d.text.strip()]
                if not valid_docs:
                    st.error("テキストを抽出できるファイルがありませんでした。")
                    st.stop()

                st.write(f"✅ {len(valid_docs)}/{len(parsed_docs)} 件の資料を正常に解析")
                status.update(label="資料解析完了", state="complete")

            # LLM分析
            with st.status(
                f"{provider} で分析中（1〜2分程度）...", expanded=True
            ) as status:
                st.write("質問リスト・論点を生成しています...")
                try:
                    new_result = analyze_documents(
                        api_key=api_key,
                        documents=valid_docs,
                        model=model,
                        provider=provider,
                    )
                except Exception as e:
                    st.error(f"API呼び出しエラー: {e}")
                    st.stop()
                status.update(label="分析完了！", state="complete")

            # スレッドとして保存
            thread_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            thread_data = _result_to_dict(new_result)
            st.session_state.threads[thread_id] = thread_data
            _save_threads(st.session_state.threads)

            st.session_state.current_thread = thread_id
            st.session_state.result = new_result
            st.rerun()

    else:
        st.markdown("---")
        st.markdown(
            """
### 📎 資料をアップロードして始めましょう

対応しているファイル形式（最大10ファイル）：
- **PDF** — IM、決算書、会社概要など
- **Excel** (.xlsx / .xls / .ods) — 財務諸表、試算表、顧客リストなど
- **Word** (.docx) — 事業計画書、組織図など
- **CSV / TSV** — 各種データファイル
- **PowerPoint** (.pptx) — プレゼン資料
- **HTML** — Webページ保存ファイル
- **テキスト / Markdown / JSON** — その他テキスト資料
"""
        )

# ---- 結果表示画面 ----
elif result:
    # 企業名 + URL
    st.markdown(f"## 📊 {result.company_name}")
    if result.company_url:
        st.markdown(
            f'<p class="company-url">🔗 <a href="{result.company_url}" target="_blank">'
            f"{result.company_url}</a></p>",
            unsafe_allow_html=True,
        )

    st.info(result.summary)

    st.markdown(
        f"**質問数**: {len(result.questions)}問 / "
        f"**主要論点**: {len(result.key_issues)}件 / "
        f"**トークン使用量**: 入力 {result.input_tokens:,} / 出力 {result.output_tokens:,}"
    )

    # Excelダウンロード
    excel_bytes = export_to_excel(result)
    fname = f"{result.company_name}_トップ面談質問_{datetime.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="📥 Excelダウンロード",
        data=excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.markdown("---")

    # タブで質問リスト・論点・財務分析を表示
    tab1, tab2, tab3 = st.tabs(["💬 質問リスト", "📈 財務分析", "⚠️ 主要論点・リスク"])

    with tab1:
        all_cats = sorted(set(q.category_name for q in result.questions))
        selected_cats = st.multiselect("カテゴリで絞り込み", all_cats, default=all_cats)
        priority_filter = st.multiselect(
            "優先度で絞り込み", ["A", "B", "C"], default=["A", "B", "C"]
        )

        filtered = [
            q
            for q in result.questions
            if q.category_name in selected_cats and q.priority in priority_filter
        ]

        priority_order = {"A": 0, "B": 1, "C": 2}
        filtered.sort(
            key=lambda q: (q.category_code, priority_order.get(q.priority, 3))
        )

        current_cat = ""
        for q in filtered:
            if q.category_name != current_cat:
                current_cat = q.category_name
                st.markdown(f"### {current_cat}")

            with st.expander(
                f"[{q.priority}] {q.question[:80]}{'...' if len(q.question) > 80 else ''}",
                expanded=(q.priority == "A"),
            ):
                st.markdown(f"**質問**: {q.question}")
                st.markdown(f"**意図**: {q.intent}")
                st.markdown(f"**背景**: {q.background}")
                if q.source_documents:
                    st.markdown(f"**関連資料**: {', '.join(q.source_documents)}")
                if q.follow_up_points:
                    st.markdown("**フォローアップ**:")
                    for fp in q.follow_up_points:
                        st.markdown(f"  - {fp}")

    with tab2:
        import pandas as pd

        fa = result.financial_analysis
        has_financial = (
            fa.pl_trends or fa.monthly_quarterly_trends or fa.bs_trends or fa.key_metrics
        )

        if not has_financial:
            st.info("財務データが資料から抽出できませんでした。財務諸表・試算表を含む資料をアップロードすると、ここに財務分析が表示されます。")
        else:
            # --- 財務コメント ---
            if fa.financial_comments:
                st.markdown("### 財務分析サマリー")
                st.info(fa.financial_comments)

            # --- 業績推移 (P/L) ---
            if fa.pl_trends:
                st.markdown("### 業績推移（P/L）")
                unit = fa.pl_trends[0].unit if fa.pl_trends else "千円"
                pl_data = []
                for p in fa.pl_trends:
                    pl_data.append({
                        "期間": p.period,
                        f"売上高（{unit}）": p.revenue,
                        f"営業利益（{unit}）": p.operating_profit,
                        f"経常利益（{unit}）": p.ordinary_profit,
                        f"当期純利益（{unit}）": p.net_income,
                        f"EBITDA（{unit}）": p.ebitda,
                    })
                df_pl = pd.DataFrame(pl_data)
                st.dataframe(df_pl, use_container_width=True, hide_index=True)

                # グラフ表示（売上高と営業利益）
                chart_cols = [c for c in df_pl.columns if c != "期間" and df_pl[c].notna().any()]
                if chart_cols:
                    df_chart = df_pl.set_index("期間")[chart_cols].apply(pd.to_numeric, errors="coerce")
                    st.bar_chart(df_chart)

            # --- 月次/四半期推移 ---
            if fa.monthly_quarterly_trends:
                trend_type = fa.monthly_quarterly_trends[0].trend_type
                label = "月次推移" if trend_type == "monthly" else "四半期推移"
                st.markdown(f"### {label}")
                unit = fa.monthly_quarterly_trends[0].unit
                mq_data = []
                for m in fa.monthly_quarterly_trends:
                    mq_data.append({
                        "期間": m.period,
                        f"売上高（{unit}）": m.revenue,
                        f"営業利益（{unit}）": m.operating_profit,
                    })
                df_mq = pd.DataFrame(mq_data)
                st.dataframe(df_mq, use_container_width=True, hide_index=True)

                chart_cols = [c for c in df_mq.columns if c != "期間" and df_mq[c].notna().any()]
                if chart_cols:
                    df_chart = df_mq.set_index("期間")[chart_cols].apply(pd.to_numeric, errors="coerce")
                    st.line_chart(df_chart)

            # --- BS推移 ---
            if fa.bs_trends:
                st.markdown("### BS推移（貸借対照表）")
                unit = fa.bs_trends[0].unit
                bs_data = []
                for b in fa.bs_trends:
                    bs_data.append({
                        "期間": b.period,
                        f"総資産（{unit}）": b.total_assets,
                        f"負債合計（{unit}）": b.total_liabilities,
                        f"純資産（{unit}）": b.net_assets,
                        f"現預金（{unit}）": b.cash_and_deposits,
                        f"有利子負債（{unit}）": b.interest_bearing_debt,
                    })
                df_bs = pd.DataFrame(bs_data)
                st.dataframe(df_bs, use_container_width=True, hide_index=True)

                chart_cols = [c for c in df_bs.columns if c != "期間" and df_bs[c].notna().any()]
                if chart_cols:
                    df_chart = df_bs.set_index("期間")[chart_cols].apply(pd.to_numeric, errors="coerce")
                    st.bar_chart(df_chart)

            # --- 主要財務指標 ---
            if fa.key_metrics:
                st.markdown("### 主要財務指標")
                for km in fa.key_metrics:
                    if km.values:
                        metric_data = []
                        for v in km.values:
                            metric_data.append({
                                "期間": v.get("period", ""),
                                km.metric_name: v.get("value", ""),
                            })
                        df_km = pd.DataFrame(metric_data)
                        st.markdown(f"**{km.metric_name}**")
                        st.dataframe(df_km, use_container_width=True, hide_index=True)

    with tab3:
        for ki in result.key_issues:
            risk_label = {"high": "🔴 高", "medium": "🟡 中", "low": "🟢 低"}.get(
                ki.risk_level, ki.risk_level
            )
            with st.expander(
                f"{risk_label} | {ki.title}",
                expanded=(ki.risk_level == "high"),
            ):
                st.markdown(ki.description)
                if ki.related_categories:
                    st.markdown(
                        f"**関連カテゴリ**: {', '.join(ki.related_categories)}"
                    )
