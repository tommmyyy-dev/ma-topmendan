"""M&A Multi-Agent Analysis — Agent Team品質をStreamlitアプリで実現"""

from __future__ import annotations

import concurrent.futures
import re
from dataclasses import dataclass


# ---------------------------------------------------------------------------
# 買い手プロファイル（デフォルト）
# ---------------------------------------------------------------------------
DEFAULT_BUYER_PROFILE = """\
## ソーシャルワイヤー株式会社（東証グロース 3929）

- 代表取締役社長: 矢田 峰之
- 親会社: 株式会社ジーニー（6562、東証プライム）※2024年7月に子会社化
- 時価総額: 約33〜37億円（2026年2月時点）
- 従業員: 約274名（連結）

### 事業セグメント
1. デジタルPRサービス（売上構成比 約59%）: @Press（プレスリリース配信）、Find Model（インフルエンサーPR）、iHack（美容特化インフルエンサーPR）
2. メディアリスニングサービス（売上構成比 約34%）: @クリッピング、ソーシャルリスニング

### 財務
| 決算期 | 売上高 | 営業利益 |
|---|---|---|
| 2024年3月期 | 36.7億円 | △0.03億円 |
| 2025年3月期 | 29.0億円 | 1.36億円 |
| 2026年3月期（予想） | 34.5億円 | 2.05億円 |

- 中期目標: 売上50億円・営業利益8億円

### M&A方針
- 目的: 既存事業の強化・シナジー創出（ボルトオン型）
- 注力領域: デジタルマーケティング、PR・広告関連
- M&A実績: @Press事業（2008年）、Find Model（2018年、2.61億円）、iHack（2025年、8.08億円）

### シナジー着眼点
1. クロスセル: @Press顧客に対象企業のサービスを販売
2. Find Model / iHack連携: インフルエンサーPR × 対象企業サービス
3. ジーニーグループシナジー: 親会社ジーニーのアドテク基盤との接続
4. のれん償却の影響: 上場企業として業績インパクトを必ず試算
5. PMIの実現可能性: iHack統合経験を踏まえた評価

### M&A担当者
- 富田 真人（iHack代表取締役社長 / M&A担当）
- 経歴: キュービック → リクルート → iHack創業
- iHack実績: 売上6.47億円、営業利益1.24億円（営利率19.2%）
"""


# ---------------------------------------------------------------------------
# Web Search (DuckDuckGo — APIキー不要)
# ---------------------------------------------------------------------------
def _web_search(query: str, max_results: int = 5) -> str:
    try:
        from ddgs import DDGS

        results = DDGS().text(query, max_results=max_results, region="jp-jp")
        if not results:
            return ""
        return "\n".join(
            f"- **{r['title']}**: {r['body']}" for r in results
        )
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# LLM API 統合呼び出し
# ---------------------------------------------------------------------------
def _call_llm(
    api_key: str,
    provider: str,
    model: str,
    system_prompt: str,
    user_prompt: str,
) -> tuple[str, int, int]:
    """(テキスト, 入力トークン, 出力トークン) を返す"""
    if "Anthropic" in provider or "Claude" in provider:
        import anthropic

        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model=model,
            max_tokens=16384,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
        return (
            resp.content[0].text,
            resp.usage.input_tokens,
            resp.usage.output_tokens,
        )
    else:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            max_tokens=16384,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        )
        return (
            resp.choices[0].message.content or "",
            resp.usage.prompt_tokens if resp.usage else 0,
            resp.usage.completion_tokens if resp.usage else 0,
        )


# ---------------------------------------------------------------------------
# Agent 1: 案件分析（IM分析 + 財務分析 + DD計画）
# ---------------------------------------------------------------------------
_MAIN_SYSTEM = """\
あなたはM&Aアドバイザリーの専門家です。バイサイド（買い手側）の立場で、
IM（企業概要書）・財務資料を分析し、包括的な案件分析レポートを作成してください。

数値は千円単位（特記がない限り）。推測は必ず「推測」と明記。
金額を文章で言及する際は億円/万円に正確に換算すること。
千円単位の数値が100,000未満なら「約X,X00万円」、100,000以上なら「約X.X億円」。
日本語で出力。"""


def _build_main_prompt(
    doc_text: str,
    financial_summary: str,
    buyer_profile: str,
) -> str:
    fin_block = ""
    if financial_summary:
        fin_block = f"""
【財務データ（CSV自動抽出・正確な数値）】
{financial_summary}
※ このデータはCSVから直接抽出した正確な数値です。分析にはこの数値を優先的に使用してください。
"""

    return f"""\
以下のIM・財務資料と買い手プロファイルを踏まえ、案件分析レポートを作成してください。

【買い手プロファイル】
{buyer_profile}

【提供資料】
{doc_text}
{fin_block}
以下の構成でMarkdownレポートを出力してください。各セクションを充実させること。

---

## 1. ディール概要
対象企業の概要（事業内容、規模、ポジション）を800〜1000字で詳述。

## 2. 対象企業の魅力（トップ5）
買い手にとっての魅力を5項目。各200字以上で具体的に。

## 3. 主要懸念事項（トップ5）
リスク・懸念を5項目。各200字以上で具体的に。

## 4. セラーの譲渡動機（仮説）
資料から読み取れる情報と一般的なM&Aの文脈から、セラーの真の譲渡動機を仮説立て。

## 5. シナジー仮説
買い手プロファイルを踏まえた具体的なシナジー仮説。クロスセル、技術連携、コスト削減、中期目標（売上50億・営利8億）への貢献度を定量的に評価。

## 6. 財務分析

### 6-1. 業績トレンド
売上高・営業利益・経常利益・当期純利益の推移をMarkdownテーブルで。トレンドの分析コメント付き。

### 6-2. 正常化EBITDA（感応度分析）
保守・ベース・楽観の3シナリオで正常化EBITDAを算出。調整項目を明記。テーブル形式。

### 6-3. バリュエーション試算
- **EV/EBITDA倍率法**: 4x / 5x / 6x / 7x / 8x の5パターンでEV・株式価値を試算。テーブル形式。
- **純資産法**: 時価純資産ベースの価値試算。
- バリュエーションレンジの総括。

### 6-4. のれん償却シミュレーション
想定取得価額でののれん額と、5年・10年・20年償却時の年間償却額を試算。買い手PLへのインパクトを評価。

### 6-5. 財務リスク（トップ5）
財務面のリスクを5項目。各200字以上。

## 7. DD計画

以下の各DD領域について、チェック項目を5〜8個ずつ列挙。
各項目に「なぜこの案件で重要か」の理由を付記。

### 7-1. ビジネスDD
### 7-2. 財務DD
### 7-3. 法務DD
### 7-4. 人事DD
### 7-5. IT DD

## 8. トップ面談で確認すべき質問（10問）
優先度の高い順に10問。各質問に質問意図と背景を付記。

---

冒頭に必ず以下を記載してください:
**対象企業**: （企業名）
"""


def _run_main(
    api_key: str, provider: str, model: str,
    doc_text: str, financial_summary: str, buyer_profile: str,
) -> tuple[str, int, int]:
    prompt = _build_main_prompt(doc_text, financial_summary, buyer_profile)
    return _call_llm(api_key, provider, model, _MAIN_SYSTEM, prompt)


# ---------------------------------------------------------------------------
# Agent 2: 市場調査（Web検索付き）
# ---------------------------------------------------------------------------
_MARKET_SYSTEM = """\
あなたはM&A案件における市場調査の専門家です。
提供されたWeb検索結果と案件情報をもとに、対象企業が属する業界の市場分析レポートを作成してください。
検索結果の情報を最大限活用し、具体的な数値・事例を含めること。
情報が不十分な場合は推測と明記した上で分析すること。日本語で出力。"""


def _build_market_prompt(
    company_name: str,
    industry: str,
    buyer_profile: str,
    search_results: str,
) -> str:
    return f"""\
以下の情報をもとに、市場調査レポートを作成してください。

【対象企業】{company_name}
【業界】{industry}

【買い手プロファイル】
{buyer_profile}

【Web検索結果】
{search_results}

以下の構成でMarkdownレポートを出力してください:

## 1. 市場規模・成長率
対象企業が属する市場の規模（金額）と成長率。関連する統計データ。

## 2. 同業界のM&A動向
直近の業界内M&A事例、買収金額、バリュエーション水準。

## 3. 競合環境
主要な競合企業、各社の特徴・ポジショニング、対象企業の位置づけ。

## 4. 業界構造変化・リスク
技術変化、規制動向、需要構造の変化など、業界が直面するリスク。

## 5. 買い手にとっての市場機会
買い手プロファイルを踏まえ、この市場に参入/拡大するメリット。
"""


def _run_market(
    api_key: str, provider: str, model: str,
    company_name: str, industry: str, buyer_profile: str,
) -> tuple[str, int, int]:
    # Web検索で市場情報を収集
    queries = [
        f"{company_name} {industry} 市場規模 成長率",
        f"{industry} M&A 動向 2025 2026",
        f"{company_name} 競合 業界",
    ]
    search_parts = []
    for q in queries:
        result = _web_search(q)
        if result:
            search_parts.append(f"**検索: {q}**\n{result}")

    search_results = "\n\n".join(search_parts) if search_parts else "(検索結果を取得できませんでした)"

    prompt = _build_market_prompt(company_name, industry, buyer_profile, search_results)
    return _call_llm(api_key, provider, model, _MARKET_SYSTEM, prompt)


# ---------------------------------------------------------------------------
# Agent 3: Red Team（批判的レビュー）
# ---------------------------------------------------------------------------
_RED_TEAM_SYSTEM = """\
あなたはM&A案件のRed Team（批判的レビュアー）です。
他のアナリストの分析結果を徹底的に批判し、見落としやリスクを指摘します。

【最重要ルール】
- 常に最大限厳しく評価する。忖度は一切不要。
- 楽観的な見通しには必ず反論する。
- 数値の前提に疑問を呈する。
- 最悪のシナリオを想定する。
日本語で出力。"""


def _build_red_team_prompt(
    main_analysis: str,
    market_research: str,
    buyer_profile: str,
) -> str:
    return f"""\
以下の案件分析と市場調査の結果を批判的にレビューしてください。

【買い手プロファイル】
{buyer_profile}

【案件分析レポート】
{main_analysis}

【市場調査レポート】
{market_research}

以下の構成でMarkdownレポートを出力してください:

## 総合判定

**判定: [GO / CONDITIONAL-GO / NO-GO]**

判定理由を3行で。

## 1. Deal Breaker候補（トップ3）
本件をNO-GOにし得る最大のリスクを3つ。各300字以上で具体的に。

## 2. 財務の疑問点
案件分析の財務データ・前提に対する疑問。数値の信頼性、会計処理の懸念、隠れた負債リスク等。

## 3. 事業リスク
ビジネスモデルの持続性、顧客集中、競合優位性の脆弱性、市場環境リスク等。

## 4. カーブアウト・譲渡リスク
親会社からの切り離しに伴うリスク。顧客・人材の離散、システム分離、ブランド毀損等。

## 5. バリュエーションの罠
楽観的なバリュエーション前提への指摘。適正倍率、比較対象の妥当性、シナジー過大評価のリスク。

## 6. ダウンサイドシナリオ（3年間PL試算）
最悪ケースでの業績推移をテーブルで試算。前提条件を明記。

| 項目 | 1年目 | 2年目 | 3年目 |
|---|---|---|---|
| 売上高 | | | |
| 営業利益 | | | |
| のれん償却後利益 | | | |

## 7. CONDITIONAL-GOの場合の条件
GOに転換するために確認・解消すべき条件を列挙。
"""


def _run_red_team(
    api_key: str, provider: str, model: str,
    main_analysis: str, market_research: str, buyer_profile: str,
) -> tuple[str, int, int]:
    prompt = _build_red_team_prompt(main_analysis, market_research, buyer_profile)
    return _call_llm(api_key, provider, model, _RED_TEAM_SYSTEM, prompt)


# ---------------------------------------------------------------------------
# 結果データクラス
# ---------------------------------------------------------------------------
@dataclass
class AgentResult:
    company_name: str
    main_analysis: str
    market_research: str
    red_team: str
    total_input_tokens: int = 0
    total_output_tokens: int = 0


# ---------------------------------------------------------------------------
# オーケストレーター
# ---------------------------------------------------------------------------
def _extract_hints(doc_text: str) -> tuple[str, str]:
    """ドキュメントから会社名・業界ヒントを抽出"""
    first = doc_text[:3000]

    # 会社名抽出
    company = ""
    for pattern in [
        r"(株式会社[^\s、。（）\n「」【】]+)",
        r"([^\s、。（）\n「」【】]+株式会社)",
        r"(有限会社[^\s、。（）\n「」【】]+)",
        r"(合同会社[^\s、。（）\n「」【】]+)",
    ]:
        m = re.search(pattern, first)
        if m:
            company = m.group(1)
            break

    # 業界ヒント
    industry = ""
    keywords = [
        "広告", "IT", "SaaS", "EC", "不動産", "人材", "製造", "建設",
        "飲食", "医療", "教育", "物流", "コンサル", "金融", "保険",
        "通信", "メディア", "ゲーム", "美容", "小売",
    ]
    for kw in keywords:
        if kw in first:
            industry = kw
            break

    return company, industry


def run_full_analysis(
    api_key: str,
    provider: str,
    model: str,
    documents: list,
    financial_summary: str = "",
    buyer_profile: str = "",
    progress_cb=None,
) -> AgentResult:
    """3エージェント並列分析を実行"""
    if not buyer_profile:
        buyer_profile = DEFAULT_BUYER_PROFILE

    # ドキュメントテキスト結合
    doc_text = "\n\n".join(
        f"===== {d.filename} ({d.doc_type}) =====\n{d.text[:30000]}"
        for d in documents
        if d.text.strip()
    )

    # 会社名・業界ヒント抽出
    company_hint, industry_hint = _extract_hints(doc_text)

    total_in = total_out = 0

    # --- Phase 1: 案件分析 + 市場調査 を並列実行 ---
    if progress_cb:
        progress_cb("🤖 案件分析 + 市場調査を並列実行中...")

    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
        f_main = ex.submit(
            _run_main, api_key, provider, model,
            doc_text, financial_summary, buyer_profile,
        )
        f_market = ex.submit(
            _run_market, api_key, provider, model,
            company_hint, industry_hint, buyer_profile,
        )

        try:
            main_text, m_in, m_out = f_main.result()
        except Exception as e:
            main_text = f"⚠️ 案件分析でエラーが発生しました: {e}"
            m_in = m_out = 0

        try:
            market_text, mk_in, mk_out = f_market.result()
        except Exception as e:
            market_text = f"⚠️ 市場調査でエラーが発生しました: {e}"
            mk_in = mk_out = 0

    total_in += m_in + mk_in
    total_out += m_out + mk_out

    # 会社名をメイン分析結果から抽出
    name_match = re.search(
        r"\*\*対象企業\*\*[：:]\s*(.+)",
        main_text[:500],
    )
    company_name = name_match.group(1).strip() if name_match else (company_hint or "不明")

    # --- Phase 2: Red Team レビュー ---
    if progress_cb:
        progress_cb("🔴 Red Team レビュー中...")

    try:
        red_text, r_in, r_out = _run_red_team(
            api_key, provider, model,
            main_text, market_text, buyer_profile,
        )
    except Exception as e:
        red_text = f"⚠️ Red Teamレビューでエラーが発生しました: {e}"
        r_in = r_out = 0

    total_in += r_in
    total_out += r_out

    return AgentResult(
        company_name=company_name,
        main_analysis=main_text,
        market_research=market_text,
        red_team=red_text,
        total_input_tokens=total_in,
        total_output_tokens=total_out,
    )
