"""
個別株式分析エンジン — 公開API（薄いオーケストレーター）

内部実装は _analyzer_*.py モジュールに分割されている:
  _analyzer_helpers.py     — safe_div, rate_change, consecutive_increase, get_latest_value
  _analyzer_thresholds.py  — DEFAULT_THRESHOLDS, INVESTOR_PROFILES, generate_dynamic_thresholds,
                              generate_evaluation_criteria
  _analyzer_quantitative.py — analyze_quantitative()
  _analyzer_screening.py   — analyze_screening()
  _analyzer_trees.py       — analyze_roa_tree(), analyze_roe_tree(), compute_pbr_contribution()

公開シンボル:
  run_full_analysis(data, benchmark=None, investor_profile='balanced') -> dict
  generate_dynamic_thresholds(benchmark, profile=None) -> dict
  INDUSTRY_LIST
"""

# ── 公開インポート ────────────────────────────────────────────────────────────
from _analyzer_thresholds import (
    DEFAULT_THRESHOLDS,
    INVESTOR_PROFILES,
    generate_dynamic_thresholds,
    generate_evaluation_criteria,
)
from _analyzer_quantitative import analyze_quantitative
from _analyzer_screening import analyze_screening
from _analyzer_trees import analyze_roa_tree, analyze_roe_tree, compute_pbr_contribution

INDUSTRY_LIST = ["製造・サービス"]

# ── データ契約 ────────────────────────────────────────────────────────────────
# parse_excel / parse_yfinance / stock_data.json が提供するキーの全一覧。
# 全パーサーはこのコントラクトに準拠すること。
#
# 配列型キー: 要素順は [最新年, 1年前, 2年前, 3年前, 4年前] (= 新しい順)
# スカラー型キー: 最新年の単一値
#
DATA_CONTRACT = {
    # ── 識別情報 ──────────────────────────────────────────────
    "company":          str,   # 会社名
    "ticker":           str,   # ティッカー記号（任意）
    "industry":         str,   # 業種（INDUSTRY_LIST 参照）

    # ── 損益計算書 ────────────────────────────────────────────
    "revenue":          list,  # 売上高 [最新..5年前]
    "eps":              list,  # 1株当たり利益 [最新..5年前]
    "op_margin":        list,  # 営業利益率(%) [最新..5年前]
    "ebitda_margin":    float, # EBITDAマージン(%) 最新値
    "ebitda_margin_5y": float, # EBITDAマージン(%) 5年前
    "cogs":             float, # 売上原価 最新値
    "cogs_5y":          float, # 売上原価 5年前
    "sga_ratio":        float, # 販管費率(%) 最新値
    "sga_ratio_5y":     float, # 販管費率(%) 5年前
    "op_income_val":    float, # 営業利益額 最新値
    "op_income_val_5y": float, # 営業利益額 5年前
    "interest_exp":     float, # 支払利息 最新値
    "interest_exp_5y":  float, # 支払利息 5年前
    "pretax_income":    float, # 税引前利益 最新値
    "pretax_income_5y": float, # 税引前利益 5年前
    "income_tax":       float, # 法人税 最新値
    "income_tax_5y":    float, # 法人税 5年前
    "net_income_val":   float, # 純利益額 最新値
    "net_income_val_5y":float, # 純利益額 5年前
    "other_exp":        float, # 営業外費用 最新値
    "other_exp_5y":     float, # 営業外費用 5年前

    # ── バランスシート ────────────────────────────────────────
    "equity_ratio":     float, # 自己資本比率(%) 最新値
    "equity_ratio_5y":  float, # 自己資本比率(%) 5年前
    "quick_ratio":      float, # 当座比率(%) 最新値
    "quick_ratio_5y":   float, # 当座比率(%) 5年前
    "current_ratio":    float, # 流動比率(%) 最新値
    "current_ratio_5y": float, # 流動比率(%) 5年前
    "total_assets":     float, # 総資産 最新値
    "total_assets_5y":  float, # 総資産 5年前
    "total_equity":     float, # 純資産 最新値
    "total_equity_5y":  float, # 純資産 5年前
    "fixed_assets":     float, # 固定資産 最新値
    "fixed_assets_5y":  float, # 固定資産 5年前
    "tangible_fixed_assets":    float,
    "tangible_fixed_assets_5y": float,
    "intangible_fixed_assets":    float,
    "intangible_fixed_assets_5y": float,
    "accounts_receivable":    float, # 売上債権 最新値
    "accounts_receivable_5y": float,
    "inventory":        float, # 棚卸資産 最新値
    "inventory_5y":     float,
    "accounts_payable": float, # 買掛金 最新値
    "accounts_payable_5y": float,

    # ── キャッシュフロー ──────────────────────────────────────
    "operating_cf":     list,  # 営業CF [最新..5年前]
    "investing_cf":     list,  # 投資CF [最新..5年前]
    "financing_cf":     list,  # 財務CF [最新..5年前]
    "fcf":              list,  # フリーCF [最新..5年前]

    # ── 収益性・効率性 ────────────────────────────────────────
    "roe":              list,  # ROE(%) [最新..5年前]
    "roe_growth_rate":  float, # ROE 5年成長率(pt/年)
    "roa":              list,  # ROA(%) [最新..5年前]
    "nopat":            float, # 税引後営業利益 最新値
    "nopat_5y":         float,
    "invested_capital": float, # 投下資本 最新値
    "invested_capital_5y": float,
    "wacc":             float, # WACC(%)

    # ── バリュエーション ──────────────────────────────────────
    "per":              float, # PER(倍) 最新値
    "per_5y":           float,
    "pbr":              float, # PBR(倍) 最新値
    "pbr_5y":           float,
    "ev":               float, # 企業価値(EV)
    "nd_ebitda":        float, # Net Debt / EBITDA
    "debt_fcf":         float, # Net Debt / FCF
    "debt_fcf_5y":      float,

    # ── 配当 ──────────────────────────────────────────────────
    "dividend_yield":   float, # 配当利回り(%) 最新値
    "dividend_yield_5y":float,
    "payout_ratio":     float, # 配当性向(%) 最新値
    "payout_ratio_5y":  float,

    # ── 定性・ESG ─────────────────────────────────────────────
    "d1_mgmt_change":   str,   # 経営陣変更 ("○"/"▲"/"×")
    "d2_ownership":     str,   # 株主構造
    "d3_esg":           str,   # ESGリスク
}


# ── バリデーション ────────────────────────────────────────────────────────────
def validate_financial_data(data: dict) -> None:
    """分析実行前にデータ品質を検証する。

    問題が見つかった場合は ValueError を raise する。
    parse_excel / parse_yfinance の出力が空だった場合に
    「解析成功 → 分析失敗」という遅延エラーを防ぐ。
    """
    if not isinstance(data, dict):
        raise ValueError("分析データは dict 型である必要があります")

    # 必須フィールドの存在確認（最低限の分析に必要なもの）
    required_fields = ["revenue", "roe", "roa"]
    missing = [f for f in required_fields if not data.get(f)]
    if missing:
        raise ValueError(
            f"必須データが不足しています: {', '.join(missing)}。"
            "Excelの列ヘッダーが認識できているか確認してください。"
        )

    # 配列型フィールドの長さ検証
    list_fields = {
        "revenue": 4,
        "eps": 4,
        "roe": 3,
        "roa": 3,
    }
    warnings = []
    for field, min_len in list_fields.items():
        val = data.get(field)
        if val is not None and isinstance(val, list):
            non_none = [v for v in val if v is not None]
            if len(non_none) < min_len:
                warnings.append(f"{field}: {len(non_none)}件（推奨{min_len}件以上）")

    # 数値範囲の基本チェック（明らかな異常値）
    equity_ratio = data.get("equity_ratio")
    if equity_ratio is not None and not (0 <= equity_ratio <= 100):
        raise ValueError(f"自己資本比率が範囲外です: {equity_ratio}（0〜100%の範囲で入力）")

    # 警告はエラーではなく data に付記（UI側でハイライト可能）
    if warnings:
        data.setdefault("_validation_warnings", []).extend(warnings)


# ── 公開エントリーポイント ────────────────────────────────────────────────────
def run_full_analysis(data: dict, benchmark=None, investor_profile: str = 'balanced') -> dict:
    """株式データの全分析を実行し、結果 dict を返す。

    Args:
        data: DATA_CONTRACT に準拠した財務データ dict（parse_excel / parse_yfinance の出力）
        benchmark: Damodaran 業界ベンチマーク dict（任意）
        investor_profile: 投資家プロファイル ID（'balanced'/'value'/'growth'/'quality'/'income'）

    Returns:
        分析結果 dict（quantitative / screening / roa_tree / roe_tree / pbr_contribution 等を含む）

    Raises:
        ValueError: データが最低限の品質要件を満たさない場合
    """
    validate_financial_data(data)

    q_results = analyze_quantitative(data, benchmark=benchmark)
    s_results = analyze_screening(data, q_results, benchmark=benchmark, investor_profile=investor_profile)
    r_results = analyze_roa_tree(data)
    roe_results = analyze_roe_tree(data)
    pbr_contrib = compute_pbr_contribution(roe_results, s_results, data, benchmark=benchmark)
    evaluation_criteria = generate_evaluation_criteria(benchmark)
    prof = INVESTOR_PROFILES.get(investor_profile, INVESTOR_PROFILES['balanced'])

    return {
        "company": data.get("company", "Unknown"),
        "ticker": data.get("ticker", ""),
        "industry": data.get("industry", "製造・サービス"),
        "investor_profile": {
            "id": investor_profile,
            "name_ja": prof["name_ja"],
            "name_en": prof["name_en"],
            "description_ja": prof["description_ja"],
            "priorities_ja": prof["priorities_ja"],
            "ref": prof["ref"],
            "weights": prof["weights"],
            "verdict": prof["verdict"],
        },
        "quantitative": q_results,
        "screening": s_results,
        "roa_tree": r_results,
        "roe_tree": roe_results,
        "pbr_contribution": pbr_contrib,
        "evaluation_criteria": evaluation_criteria,
        "validation_warnings": data.get("_validation_warnings", []),
        "raw_data": {
            "revenue": data.get("revenue"),
            "fcf": data.get("fcf"),
            "eps": data.get("eps"),
            "roe": data.get("roe"),
            "roa": data.get("roa"),
            "op_margin": data.get("op_margin"),
            "operating_cf": data.get("operating_cf"),
            "investing_cf": data.get("investing_cf"),
            "financing_cf": data.get("financing_cf"),
            "equity_ratio": data.get("equity_ratio"),
            "equity_ratio_5y": data.get("equity_ratio_5y"),
            "quick_ratio": data.get("quick_ratio"),
            "quick_ratio_5y": data.get("quick_ratio_5y"),
            "current_ratio": data.get("current_ratio"),
            "current_ratio_5y": data.get("current_ratio_5y"),
            "debt_fcf": data.get("debt_fcf"),
            "debt_fcf_5y": data.get("debt_fcf_5y"),
            "ebitda_margin": data.get("ebitda_margin"),
            "ebitda_margin_5y": data.get("ebitda_margin_5y"),
        },
    }
