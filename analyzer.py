"""LLM連携: Claude API / OpenAI API を使った資料分析・質問生成"""

from __future__ import annotations

import json
import urllib.parse
import urllib.request
from dataclasses import dataclass, field

from parsers import ParsedDocument

# ---------------------------------------------------------------------------
# カテゴリ定義
# ---------------------------------------------------------------------------
CATEGORIES = {
    "BIZ": "事業概要・ビジネスモデル",
    "MKT": "市場環境・競合",
    "FIN": "財務分析・収益性",
    "ORG": "組織・人事・キーパーソン",
    "CST": "顧客・取引先",
    "TEC": "技術・知的財産",
    "RSK": "リスク・法務・コンプライアンス",
    "GRW": "成長戦略・事業計画",
    "SYN": "シナジー・PMI",
    "GOV": "株主・ガバナンス",
    "TRN": "譲渡理由・条件",
}

# ---------------------------------------------------------------------------
# システムプロンプト
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """\
あなたはM&A（企業買収）のバイサイドにおけるデューデリジェンス及びトップ面談（経営者面談）の\
専門家アドバイザーです。上場企業のM&A担当者が対象企業のトップマネジメントとの面談で\
確認すべき質問事項・論点を策定する業務を支援します。

以下の原則に従ってください：
1. 質問は具体的かつ実務的であること（抽象的な質問は避ける）
2. 各質問には、なぜその質問が重要なのか（質問意図）を明記すること
3. 提供された資料から読み取れる情報に基づき、矛盾点・不明点・深掘りすべき論点を特定すること
4. 資料に記載がない重要事項についても、M&Aデューデリジェンスの観点から質問を提案すること
5. 質問は日本語のビジネス敬語で、トップマネジメントに対して失礼のない表現を使うこと
6. 数値やファクトに基づく質問を優先し、感想や主観的な質問は避けること
"""

# ---------------------------------------------------------------------------
# 分析プロンプト
# ---------------------------------------------------------------------------
def _build_analysis_prompt(
    documents: list[ParsedDocument],
    max_chars_per_doc: int = 30000,
) -> str:
    category_list = "\n".join(
        f"  - {code}: {name}" for code, name in CATEGORIES.items()
    )

    doc_sections: list[str] = []
    for doc in documents:
        if not doc.text.strip():
            continue
        text = doc.text[:max_chars_per_doc]
        if len(doc.text) > max_chars_per_doc:
            text += "\n\n... (以降省略)"
        doc_sections.append(
            f"===== {doc.filename} ({doc.doc_type}) =====\n{text}"
        )

    all_docs_text = "\n\n".join(doc_sections)

    return f"""\
以下はM&A案件における対象企業の提供資料です。これらを総合的に分析し、
トップ面談（経営者面談）用の質問リスト・論点を作成してください。

【提供資料】
{all_docs_text}

【指示】
上記の資料を分析し、以下の3つのセクションで出力してください。

■ セクション1: 企業サマリー
対象企業の概要を400字程度で簡潔にまとめてください。
事業内容、財務状況の要点、特徴的な点を含めてください。

■ セクション2: 質問リスト
以下のカテゴリ別に質問を作成してください（カテゴリあたり2〜6問目安）。
{category_list}

■ セクション3: 主要論点・リスク
資料を横断的に分析し、ディールの成否に関わる主要な論点・リスクを5〜10点挙げてください。

【出力フォーマット】
必ず以下のJSON形式で出力してください。JSON以外のテキストは含めないでください。

{{
  "company_name": "対象企業名（資料から推定）",
  "summary": "企業サマリー（400字程度）",
  "questions": [
    {{
      "category_code": "BIZ",
      "category_name": "事業概要・ビジネスモデル",
      "question": "質問文（丁寧語）",
      "intent": "質問意図（なぜ確認が必要か）",
      "background": "質問の背景（資料のどの情報から導かれたか）",
      "source_documents": ["関連する資料名"],
      "priority": "A",
      "follow_up_points": ["回答次第で追加確認すべき事項"]
    }}
  ],
  "key_issues": [
    {{
      "title": "論点タイトル",
      "description": "論点の詳細説明",
      "risk_level": "high",
      "related_categories": ["FIN", "RSK"]
    }}
  ]
}}

priorityは以下の基準で付与してください：
  A: 必須確認（ディールの判断に直結）
  B: 重要（バリュエーションや統合に影響）
  C: できれば確認（補足的な情報）

risk_levelは high / medium / low から選択してください。
"""


# ---------------------------------------------------------------------------
# データクラス
# ---------------------------------------------------------------------------
@dataclass
class Question:
    category_code: str
    category_name: str
    question: str
    intent: str
    background: str
    source_documents: list[str]
    priority: str  # A / B / C
    follow_up_points: list[str]


@dataclass
class KeyIssue:
    title: str
    description: str
    risk_level: str
    related_categories: list[str]


@dataclass
class AnalysisResult:
    company_name: str
    summary: str
    questions: list[Question]
    key_issues: list[KeyIssue]
    raw_json: dict = field(default_factory=dict)
    input_tokens: int = 0
    output_tokens: int = 0
    company_url: str = ""


# ---------------------------------------------------------------------------
# JSON抽出ヘルパー
# ---------------------------------------------------------------------------
def _extract_json(raw_text: str) -> dict:
    """LLMレスポンスからJSONを抽出する"""
    text = raw_text.strip()
    if "```json" in text:
        text = text.split("```json")[1].split("```")[0].strip()
    elif "```" in text:
        text = text.split("```")[1].split("```")[0].strip()
    return json.loads(text)


def _parse_result(data: dict, input_tokens: int, output_tokens: int) -> AnalysisResult:
    """JSONデータをAnalysisResultに変換する"""
    questions = [
        Question(
            category_code=q.get("category_code", ""),
            category_name=q.get("category_name", ""),
            question=q.get("question", ""),
            intent=q.get("intent", ""),
            background=q.get("background", ""),
            source_documents=q.get("source_documents", []),
            priority=q.get("priority", "C"),
            follow_up_points=q.get("follow_up_points", []),
        )
        for q in data.get("questions", [])
    ]

    key_issues = [
        KeyIssue(
            title=ki.get("title", ""),
            description=ki.get("description", ""),
            risk_level=ki.get("risk_level", "medium"),
            related_categories=ki.get("related_categories", []),
        )
        for ki in data.get("key_issues", [])
    ]

    return AnalysisResult(
        company_name=data.get("company_name", "（不明）"),
        summary=data.get("summary", ""),
        questions=questions,
        key_issues=key_issues,
        raw_json=data,
        input_tokens=input_tokens,
        output_tokens=output_tokens,
    )


# ---------------------------------------------------------------------------
# Anthropic (Claude) API
# ---------------------------------------------------------------------------
def _call_anthropic(
    api_key: str,
    documents: list[ParsedDocument],
    model: str,
) -> AnalysisResult:
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    user_prompt = _build_analysis_prompt(documents)

    response = client.messages.create(
        model=model,
        max_tokens=8192,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_prompt}],
    )

    raw_text = response.content[0].text.strip()
    in_tok = response.usage.input_tokens
    out_tok = response.usage.output_tokens

    try:
        data = _extract_json(raw_text)
    except json.JSONDecodeError:
        return AnalysisResult(
            company_name="（解析エラー）",
            summary=raw_text[:500],
            questions=[], key_issues=[], raw_json={},
            input_tokens=in_tok, output_tokens=out_tok,
        )

    return _parse_result(data, in_tok, out_tok)


# ---------------------------------------------------------------------------
# OpenAI (GPT) API
# ---------------------------------------------------------------------------
def _call_openai(
    api_key: str,
    documents: list[ParsedDocument],
    model: str,
) -> AnalysisResult:
    from openai import OpenAI, RateLimitError

    client = OpenAI(api_key=api_key)

    # トークン上限エラー時に自動でテキスト量を減らしてリトライ
    char_limits = [30000, 15000, 8000, 4000]
    last_error = None

    for limit in char_limits:
        user_prompt = _build_analysis_prompt(documents, max_chars_per_doc=limit)
        try:
            response = client.chat.completions.create(
                model=model,
                max_tokens=8192,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                response_format={"type": "json_object"},
            )
            break
        except RateLimitError as e:
            last_error = e
            continue
    else:
        raise last_error  # type: ignore[misc]

    raw_text = response.choices[0].message.content.strip()
    in_tok = response.usage.prompt_tokens if response.usage else 0
    out_tok = response.usage.completion_tokens if response.usage else 0

    try:
        data = _extract_json(raw_text)
    except json.JSONDecodeError:
        return AnalysisResult(
            company_name="（解析エラー）",
            summary=raw_text[:500],
            questions=[], key_issues=[], raw_json={},
            input_tokens=in_tok, output_tokens=out_tok,
        )

    return _parse_result(data, in_tok, out_tok)


# ---------------------------------------------------------------------------
# 統合エントリーポイント
# ---------------------------------------------------------------------------
# プロバイダー別モデル一覧
PROVIDERS = {
    "Anthropic (Claude)": {
        "models": ["claude-sonnet-4-20250514", "claude-haiku-4-5-20251001"],
        "env_key": "ANTHROPIC_API_KEY",
    },
    "OpenAI (GPT)": {
        "models": ["gpt-4o", "gpt-4o-mini"],
        "env_key": "OPENAI_API_KEY",
    },
}


def analyze_documents(
    api_key: str,
    documents: list[ParsedDocument],
    model: str = "gpt-4o",
    provider: str = "OpenAI (GPT)",
) -> AnalysisResult:
    """ドキュメントを分析してトップ面談用の質問リストを生成する"""
    if provider == "Anthropic (Claude)":
        result = _call_anthropic(api_key, documents, model)
    else:
        result = _call_openai(api_key, documents, model)

    # 会社URLを検索して付与
    if result.company_name and result.company_name not in ("（不明）", "（解析エラー）"):
        result.company_url = search_company_url(result.company_name)

    return result


# ---------------------------------------------------------------------------
# 会社URL検索
# ---------------------------------------------------------------------------
def search_company_url(company_name: str) -> str:
    """会社名からGoogle検索し、公式サイトURLを推定して返す"""
    try:
        query = urllib.parse.quote(f"{company_name} 公式サイト")
        url = f"https://www.google.com/search?q={query}&num=5"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
        # 検索結果からURLを抽出（/url?q=...&の形式）
        import re
        urls = re.findall(r'/url\?q=(https?://[^&"]+)', html)
        # google自身やキャッシュを除外
        for u in urls:
            decoded = urllib.parse.unquote(u)
            if "google." in decoded or "youtube." in decoded or "wikipedia." in decoded:
                continue
            return decoded
    except Exception:
        pass
    return ""
