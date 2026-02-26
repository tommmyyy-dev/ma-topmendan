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
from parsers import (
    SUPPORTED_EXTENSIONS,
    ExtractedFinancials,
    ParsedDocument,
    apply_period_labels_from_llm,
    extract_financial_data,
    extract_financial_summary,
    parse_file,
    PL_DISPLAY_NAMES,
    BS_DISPLAY_NAMES,
)

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
                "ordinary_profit": m.ordinary_profit,
                "net_income": m.net_income,
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


def _result_to_dict(
    r: AnalysisResult,
    extracted_financials: ExtractedFinancials | None = None,
) -> dict:
    d = {
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
                "ma_risk_implications": ki.ma_risk_implications,
                "post_merger_solutions": ki.post_merger_solutions,
            }
            for ki in r.key_issues
        ],
        "financial_analysis": _financial_to_dict(r.financial_analysis),
        "input_tokens": r.input_tokens,
        "output_tokens": r.output_tokens,
        "created_at": datetime.now().isoformat(),
    }
    if extracted_financials:
        d["extracted_financials"] = extracted_financials.to_dict()
    return d


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
                ordinary_profit=m.get("ordinary_profit"),
                net_income=m.get("net_income"),
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
                ma_risk_implications=ki.get("ma_risk_implications", ""),
                post_merger_solutions=ki.get("post_merger_solutions", ""),
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
if "extracted_financials" not in st.session_state:
    st.session_state.extracted_financials = None

# ---------------------------------------------------------------------------
# サイドバー
# ---------------------------------------------------------------------------
with st.sidebar:
    # ---- 新規作成ボタン ----
    if st.button("＋ 新しい案件の分析を作成", use_container_width=True, type="primary"):
        st.session_state.current_thread = None
        st.session_state.result = None
        st.session_state.extracted_financials = None
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
                ef_data = t.get("extracted_financials")
                st.session_state.extracted_financials = (
                    ExtractedFinancials.from_dict(ef_data) if ef_data else None
                )
                st.rerun()

            if is_active:
                # 削除ボタン
                if st.button("🗑 この案件を削除", key=f"del_{tid}", type="secondary"):
                    del st.session_state.threads[tid]
                    _save_threads(st.session_state.threads)
                    st.session_state.current_thread = None
                    st.session_state.result = None
                    st.session_state.extracted_financials = None
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

    api_key = st.text_input(
        f"{provider} API Key",
        type="password",
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
                raw_files: list[tuple[str, bytes]] = []
                for f in uploaded_files:
                    st.write(f"📄 {f.name} を解析中...")
                    file_bytes = f.read()
                    raw_files.append((f.name, file_bytes))
                    doc = parse_file(f.name, file_bytes)
                    parsed_docs.append(doc)
                    if doc.warnings:
                        for w in doc.warnings:
                            st.warning(f"  ⚠ {f.name}: {w}")

                valid_docs = [d for d in parsed_docs if d.text.strip()]
                if not valid_docs:
                    st.error("テキストを抽出できるファイルがありませんでした。")
                    st.stop()

                # 財務CSVから正確なデータを事前抽出
                fin_data = extract_financial_data(raw_files)
                fin_summary = extract_financial_summary(raw_files)
                if fin_data.pl_list or fin_data.bs_list:
                    st.write(f"📊 財務CSVから P/L {len(fin_data.pl_list)}期分、B/S {len(fin_data.bs_list)}期分を直接抽出しました")

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
                        financial_summary=fin_summary,
                    )
                except Exception as e:
                    st.error(f"API呼び出しエラー: {e}")
                    st.stop()
                status.update(label="分析完了！", state="complete")

            # LLMの年度ラベルをCSV抽出データに適用
            if fin_data.pl_list or fin_data.bs_list:
                llm_fa = new_result.raw_json.get("financial_analysis", {})
                apply_period_labels_from_llm(
                    fin_data,
                    llm_fa.get("pl_trends", []),
                    llm_fa.get("bs_trends", []),
                )

            # スレッドとして保存
            thread_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            thread_data = _result_to_dict(new_result, fin_data)
            st.session_state.threads[thread_id] = thread_data
            _save_threads(st.session_state.threads)

            st.session_state.current_thread = thread_id
            st.session_state.result = new_result
            st.session_state.extracted_financials = fin_data
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
    excel_bytes = export_to_excel(result, st.session_state.get("extracted_financials"))
    fname = f"{result.company_name}_トップ面談質問_{datetime.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="📥 Excelダウンロード",
        data=excel_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.markdown("---")

    # タブで質問リスト・論点・財務分析・企業サマリーを表示
    tab1, tab2, tab3, tab4 = st.tabs(["💬 質問リスト", "📈 財務分析", "⚠️ 主要論点・リスク", "🏢 企業サマリー"])

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
            key=lambda q: (priority_order.get(q.priority, 3), q.category_code)
        )

        priority_labels = {"A": "🔴 優先度A（必須確認）", "B": "🟡 優先度B（重要）", "C": "🟢 優先度C（補足）"}
        current_priority = ""
        for i, q in enumerate(filtered, 1):
            if q.priority != current_priority:
                current_priority = q.priority
                st.markdown(f"### {priority_labels.get(q.priority, q.priority)}")

            with st.expander(
                f"Q{i}. [{q.category_code}] {q.question[:80]}{'...' if len(q.question) > 80 else ''}",
                expanded=(q.priority == "A"),
            ):
                st.markdown(f"**カテゴリ**: {q.category_name}")
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

        def _fmt_number_cols(df: pd.DataFrame) -> pd.DataFrame:
            """数値列をカンマ区切りフォーマットした文字列に変換"""
            df_display = df.copy()
            for col in df_display.columns:
                if col == "期間":
                    continue
                try:
                    numeric_vals = pd.to_numeric(df_display[col], errors="coerce")
                    df_display[col] = numeric_vals.apply(
                        lambda x: f"{int(x):,}" if pd.notna(x) else "-"
                    )
                except Exception:
                    pass
            return df_display

        fa = result.financial_analysis
        ef: ExtractedFinancials | None = st.session_state.get("extracted_financials")

        has_csv_data = ef and (ef.pl_list or ef.bs_list)
        has_llm_financial = fa.key_metrics or fa.financial_comments

        if not has_csv_data and not has_llm_financial:
            st.info("財務データが資料から抽出できませんでした。財務諸表・試算表を含む資料をアップロードすると、ここに財務分析が表示されます。")
        else:
            # --- 財務コメント (LLM) ---
            if fa.financial_comments:
                st.markdown("### 財務分析サマリー")
                st.info(fa.financial_comments)

            # --- 業績推移 (P/L) - CSV直接データ ---
            if has_csv_data and ef.pl_list:
                st.markdown("### 業績推移（P/L）")
                st.caption(f"CSVから直接抽出（{len(ef.pl_list)}期分）・単位: 円")
                pl_data = []
                for pl in ef.pl_list:
                    row: dict = {"期間": pl.period_label}
                    row["売上高"] = pl.revenue
                    row["売上原価"] = pl.cost_of_sales
                    row["売上総利益"] = pl.gross_profit
                    row["販管費"] = pl.sga
                    row["営業利益"] = pl.operating_profit
                    row["経常利益"] = pl.ordinary_profit
                    row["当期純利益"] = pl.net_income
                    pl_data.append(row)
                df_pl = pd.DataFrame(pl_data)
                st.dataframe(_fmt_number_cols(df_pl), use_container_width=True, hide_index=True)

                # 売上高グラフ
                if df_pl["売上高"].notna().any():
                    st.markdown("**売上高推移**")
                    df_rev = df_pl[["期間", "売上高"]].copy()
                    df_rev["売上高"] = pd.to_numeric(df_rev["売上高"], errors="coerce")
                    st.bar_chart(df_rev.set_index("期間"))

                # 営業利益グラフ
                if df_pl["営業利益"].notna().any():
                    st.markdown("**営業利益推移**")
                    df_op = df_pl[["期間", "営業利益"]].copy()
                    df_op["営業利益"] = pd.to_numeric(df_op["営業利益"], errors="coerce")
                    st.bar_chart(df_op.set_index("期間"))

            # --- BS推移 - CSV直接データ ---
            if has_csv_data and ef.bs_list:
                st.markdown("### BS推移（貸借対照表）")
                st.caption(f"CSVから直接抽出（{len(ef.bs_list)}期分）・単位: 円")
                bs_data = []
                for bs in ef.bs_list:
                    row: dict = {"期間": bs.period_label}
                    row["総資産"] = bs.total_assets
                    row["負債合計"] = bs.total_liabilities
                    row["純資産"] = bs.net_assets
                    row["現預金"] = bs.cash
                    row["有利子負債"] = bs.interest_bearing_debt
                    bs_data.append(row)
                df_bs = pd.DataFrame(bs_data)
                st.dataframe(_fmt_number_cols(df_bs), use_container_width=True, hide_index=True)

            # --- 主要財務指標 (LLM) ---
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
        st.markdown(f"**{len(result.key_issues)}件** の主要論点・リスクを特定")
        for i, ki in enumerate(result.key_issues, 1):
            risk_label = {"high": "🔴 高", "medium": "🟡 中", "low": "🟢 低"}.get(
                ki.risk_level, ki.risk_level
            )
            with st.expander(
                f"{risk_label} | {i}. {ki.title}",
                expanded=(ki.risk_level == "high"),
            ):
                st.markdown("**概要・詳細**")
                st.markdown(ki.description)
                if ki.ma_risk_implications:
                    st.markdown("---")
                    st.markdown("**M&Aリスクへの影響**")
                    st.markdown(ki.ma_risk_implications)
                if ki.post_merger_solutions:
                    st.markdown("---")
                    st.markdown("**買収後の対策・解決策**")
                    st.markdown(ki.post_merger_solutions)
                if ki.related_categories:
                    st.markdown("---")
                    st.markdown(
                        f"**関連カテゴリ**: {', '.join(ki.related_categories)}"
                    )

    with tab4:
        st.markdown("### 企業サマリー")
        st.markdown(result.summary)
        if result.company_url:
            st.markdown(f"**企業URL**: [{result.company_url}]({result.company_url})")
