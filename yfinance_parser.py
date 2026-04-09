"""
yfinanceからティッカーシンボルで財務データを取得し、
parse_excel()と同一形式の (data, ts_data) を返すパーサー
"""
import logging
import math
import threading
import time
import yfinance as yf

logger = logging.getLogger(__name__)

# ── モジュールレベルキャッシュ（^TNX リスクフリーレート）─────────────────────
# CPython の GIL で単純な代入はスレッドセーフだが、check-then-act を保護するため Lock を使用
_tnx_cache: dict = {"rate": None, "ts": 0.0}
_tnx_lock = threading.Lock()
_TNX_CACHE_TTL: float = 86400.0  # 24時間


# yfinance行名 → 内部キーのマッピング
_INCOME_MAP = {
    'Total Revenue': 'revenue',
    'Operating Revenue': 'revenue',
    'Cost Of Revenue': 'cogs',
    'Reconciled Cost Of Revenue': 'cogs',
    'Operating Income': 'op_income',
    'Net Income': 'net_income',
    'Net Income Common Stockholders': 'net_income',
    'Basic EPS': 'eps',
    'Diluted EPS': 'eps_diluted',
    'Selling General And Administration': 'sga',
    'Interest Expense Non Operating': 'interest_exp',
    'Interest Expense': 'interest_exp',
    'Net Interest Income': 'net_interest',
    'Other Income Expense': 'other_exp',
    'Other Non Operating Income Expenses': 'other_exp',
    'Pretax Income': 'pretax_income',
    'Tax Provision': 'income_tax',
    'EBITDA': 'ebitda',
    'Gross Profit': 'gross_profit',
    'Reconciled Depreciation': 'da',
}

_CASHFLOW_MAP = {
    'Free Cash Flow': 'fcf',
    'Operating Cash Flow': 'ocf',
    'Cash Flow From Continuing Operating Activities': 'ocf',
    'Capital Expenditure': 'capex',
    'Investing Cash Flow': 'investing_cf',
    'Cash Flow From Continuing Investing Activities': 'investing_cf',
    'Financing Cash Flow': 'financing_cf',
    'Cash Flow From Continuing Financing Activities': 'financing_cf',
    'Depreciation And Amortization': 'da',
    'Depreciation Amortization Depletion': 'da',
}

_BALANCE_MAP = {
    'Total Assets': 'total_assets',
    'Stockholders Equity': 'total_equity',
    'Common Stock Equity': 'total_equity',
    'Total Debt': 'total_debt',
    'Accounts Receivable': 'receivables',
    'Receivables': 'receivables',
    'Inventory': 'inventory',
    'Accounts Payable': 'payables',
    'Payables': 'payables',
    'Current Assets': 'current_assets',
    'Current Liabilities': 'current_liab',
    'Cash And Cash Equivalents': 'cash',
    'Cash Cash Equivalents And Short Term Investments': 'cash_and_st',
    'Net PPE': 'fixed_assets',
    'Goodwill And Other Intangible Assets': 'intangibles',
    'Other Intangible Assets': 'intangibles_other',
    'Net Debt': 'net_debt',
    'Total Non Current Assets': 'long_term_assets',
    'Long Term Debt': 'long_term_debt',
    'Retained Earnings': 'retained_earnings',
    'Invested Capital': 'invested_capital',
    'Working Capital': 'working_capital',
    'Tangible Book Value': 'tangible_book',
}


def _safe(v):
    """NaN/Inf → None 変換"""
    if v is None:
        return None
    try:
        if math.isnan(v) or math.isinf(v):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def _extract_series(df, mapping):
    """DataFrameから指定マッピングに従い {key: [新しい→古い順のリスト]} を返す。
    yfinanceのDataFrameは列が日付（新しい順）、行が項目名。"""
    result = {}
    if df is None or df.empty:
        return result, []

    # 日付（列名）を文字列に変換
    dates = [str(c)[:4] for c in df.columns]

    for row_name, key in mapping.items():
        if row_name in df.index:
            vals = [_safe(df.loc[row_name].iloc[i]) for i in range(len(df.columns))]
            # 重複キーは最初に見つかった（より優先度の高い）ものを保持
            if key not in result:
                result[key] = vals
            else:
                # 既存データが全Noneなら上書き
                if all(v is None for v in result[key]):
                    result[key] = vals

    return result, dates


def _get_year_end_price(hist_df, year_str, is_japan=False):
    """株価履歴から当該年度末の終値を返す。"""
    if hist_df is None or hist_df.empty:
        return None
    try:
        year = int(year_str)
        # 日本株: 3月決算が主流 → 3月末
        # 米国株: 12月決算が主流 → 12月末
        target_month = 3 if is_japan else 12
        target_year = year

        mask = (hist_df.index.year == target_year) & (hist_df.index.month == target_month)
        prices = hist_df.loc[mask, 'Close']
        if len(prices) > 0:
            return float(prices.iloc[-1])

        # 対象月のデータがなければ前後1ヶ月を探す
        for delta in [1, -1, 2, -2]:
            m = target_month + delta
            y = target_year
            if m > 12:
                m -= 12
                y += 1
            elif m < 1:
                m += 12
                y -= 1
            mask2 = (hist_df.index.year == y) & (hist_df.index.month == m)
            prices2 = hist_df.loc[mask2, 'Close']
            if len(prices2) > 0:
                return float(prices2.iloc[-1])
    except Exception as e:
        logger.warning("_get_year_end_price: 株価取得失敗 (year=%s, month=%s): %s", target_year, target_month, e)
    return None


def _assess_esg(sustainability_df):
    """ESGデータから定性スコアを返す ("○"/"▲"/"×")。"""
    if sustainability_df is None or sustainability_df.empty:
        return "○"
    try:
        val = sustainability_df.loc['totalEsg', 'Value'] if 'totalEsg' in sustainability_df.index else None
        if val is None:
            # カラムが違う形式の場合
            if 'Value' not in sustainability_df.columns and len(sustainability_df.columns) > 0:
                col = sustainability_df.columns[0]
                val = sustainability_df['totalEsg'][col] if 'totalEsg' in sustainability_df.index else None
        if val is not None:
            score = float(val)
            # ESGリスクスコア: 低いほど良い (0-10: negligible, 10-20: low, 20-30: medium, 30+: high)
            if score < 20:
                return "○"
            elif score < 30:
                return "▲"
            else:
                return "×"
    except Exception as e:
        logger.warning("_assess_esg: ESGスコア評価失敗: %s", e)
    return "○"


def _assess_ownership(major_holders_df):
    """機関投資家保有比率から株主構造スコアを返す ("○"/"▲"/"×")。"""
    if major_holders_df is None or major_holders_df.empty:
        return "○"
    try:
        # major_holders の行構成: 0=insider%, 1=institution%, 2=float%, 3=outstanding%
        inst_pct = None
        insider_pct = None
        for idx in major_holders_df.index:
            row_label = str(major_holders_df.iloc[idx, 1]).lower() if major_holders_df.shape[1] > 1 else ""
            val_str = str(major_holders_df.iloc[idx, 0])
            val = float(val_str.replace('%', '')) / 100 if '%' in val_str else float(val_str)
            if 'institution' in row_label:
                inst_pct = val
            elif 'insider' in row_label:
                insider_pct = val

        if inst_pct is None:
            # フォールバック: 最初の数値行
            try:
                val_str = str(major_holders_df.iloc[1, 0])
                val = float(val_str.replace('%', '')) / 100 if '%' in val_str else float(val_str)
                inst_pct = val
            except (IndexError, ValueError, TypeError) as e:
                logger.warning("_assess_ownership: フォールバック行取得失敗: %s", e)

        if inst_pct is not None:
            # 機関投資家比率が高い = 透明性が高い・ガバナンス良好
            if inst_pct >= 0.50:
                return "○"
            elif inst_pct >= 0.20:
                return "▲"
            else:
                return "×"
    except Exception as e:
        logger.warning("_assess_ownership: 株主構造評価失敗: %s", e)
    return "○"


def _get_analyst_consensus(recommendations_df):
    """アナリスト推奨から合意評価を返す。"""
    if recommendations_df is None or recommendations_df.empty:
        return None, None
    try:
        # 直近30件を対象
        recent = recommendations_df.tail(30)
        grade_col = None
        for col in ['To Grade', 'toGrade', 'Action']:
            if col in recent.columns:
                grade_col = col
                break
        if grade_col is None:
            return None, None

        grades = recent[grade_col].dropna().str.lower()
        buy_kw = ['buy', 'outperform', 'overweight', 'strong buy', 'accumulate', 'add']
        sell_kw = ['sell', 'underperform', 'underweight', 'strong sell', 'reduce']
        hold_kw = ['hold', 'neutral', 'equal weight', 'market perform', 'in-line']

        buys = sum(1 for g in grades if any(k in g for k in buy_kw))
        sells = sum(1 for g in grades if any(k in g for k in sell_kw))
        holds = sum(1 for g in grades if any(k in g for k in hold_kw))
        total = buys + sells + holds
        if total == 0:
            return None, None

        buy_pct = buys / total
        sell_pct = sells / total
        if buy_pct >= 0.60:
            consensus = "Strong Buy"
        elif buy_pct >= 0.40:
            consensus = "Buy"
        elif sell_pct >= 0.40:
            consensus = "Sell"
        else:
            consensus = "Hold"

        return consensus, {"buy": buys, "hold": holds, "sell": sells}
    except Exception as e:
        logger.warning("_get_analyst_consensus: アナリスト推奨集計失敗: %s", e)
        return None, None


def _get_risk_free_rate(is_jpy: bool = False) -> float:
    """リスクフリーレートを返す。米国株は^TNX（24hキャッシュ）、日本株はJGB~1%を使用。"""
    if is_jpy:
        # 日本株: 日本国債10年利回り (JGB ~1%)
        return 0.010

    global _tnx_cache
    now = time.time()
    if _tnx_cache["rate"] is not None and (now - _tnx_cache["ts"]) < _TNX_CACHE_TTL:
        return _tnx_cache["rate"]

    rate = 0.045  # フォールバック: 4.5%
    try:
        tnx = yf.Ticker('^TNX')
        tnx_price = _safe(tnx.info.get('regularMarketPrice'))
        if tnx_price and 0 < tnx_price < 20:
            rate = tnx_price / 100
    except Exception as e:
        logger.warning("_get_risk_free_rate: ^TNX 取得失敗、フォールバック 4.5%% を使用: %s", e)

    _tnx_cache["rate"] = rate
    _tnx_cache["ts"] = now
    return rate


def _calc_dividend_growth_rate(dividends_series) -> float | None:
    """配当履歴から5年間の配当成長率(CAGR%)を計算する。"""
    if dividends_series is None or len(dividends_series) < 2:
        return None
    try:
        # 年次配当合計で計算（直近5年分）
        by_year = dividends_series.groupby(dividends_series.index.year).sum()
        by_year = by_year[by_year > 0]
        if len(by_year) < 2:
            return None

        years = sorted(by_year.index)
        # 利用可能な最大5年スパン
        span = min(5, len(years) - 1)
        if span < 1:
            return None

        d_start = float(by_year.loc[years[-(span + 1)]])
        d_end = float(by_year.loc[years[-1]])
        if d_start <= 0 or d_end <= 0:
            return None

        cagr = ((d_end / d_start) ** (1.0 / span) - 1.0) * 100
        # 極端な値はNone (例: -100% ~ +100%)
        return round(cagr, 2) if -50 < cagr < 100 else None
    except Exception:
        return None


def parse_yfinance(ticker_symbol):
    """yfinanceからデータを取得し、parse_excel()と同一形式の(data, ts_data)を返す。

    Args:
        ticker_symbol: ティッカーシンボル (例: '7203.T', 'AAPL')

    Returns:
        (data, ts_data) タプル。parse_excel()と同一のキー構造。

    Raises:
        ValueError: ティッカーが無効またはデータ取得失敗時
    """
    ticker = yf.Ticker(ticker_symbol)

    # ── 基本財務データ取得 ────────────────────────────────────────────────────
    try:
        inc_df = ticker.financials
    except Exception:
        inc_df = None
    try:
        bs_df = ticker.balance_sheet
    except Exception:
        bs_df = None
    try:
        cf_df = ticker.cashflow
    except Exception:
        cf_df = None
    try:
        info = ticker.info or {}
    except Exception:
        info = {}

    if (inc_df is None or inc_df.empty) and (bs_df is None or bs_df.empty):
        raise ValueError(f"ティッカー '{ticker_symbol}' のデータを取得できませんでした。シンボルを確認してください。")

    # ── 追加データ取得（エラーは無視して続行）────────────────────────────────
    try:
        hist_df = ticker.history(period='6y', interval='1mo')
    except Exception:
        hist_df = None

    try:
        sustainability_df = ticker.sustainability
    except Exception:
        sustainability_df = None

    try:
        recommendations_df = ticker.recommendations
    except Exception:
        recommendations_df = None

    try:
        major_holders_df = ticker.major_holders
    except Exception:
        major_holders_df = None

    try:
        dividends_series = ticker.dividends
    except Exception:
        dividends_series = None

    # ── 財務データ抽出 ────────────────────────────────────────────────────────
    inc_data, inc_dates = _extract_series(inc_df, _INCOME_MAP)
    cf_data, cf_dates = _extract_series(cf_df, _CASHFLOW_MAP)
    bs_data, bs_dates = _extract_series(bs_df, _BALANCE_MAP)

    dates = inc_dates or cf_dates or bs_dates
    all_data = {}
    all_data.update(bs_data)
    all_data.update(cf_data)
    all_data.update(inc_data)  # incが最優先

    def g(key, idx=0):
        lst = all_data.get(key, [])
        return lst[idx] if idx < len(lst) else None

    def g_list(key):
        return all_data.get(key, [])

    n = len(dates)

    # ── 通貨・地域メタデータ（早期決定） ─────────────────────────────────────
    currency = info.get("currency", "USD")
    is_jpy = currency == "JPY" or ticker_symbol.endswith(".T") or ticker_symbol.endswith(".J")

    # ── 実効税率（年別）──────────────────────────────────────────────────────
    # income_tax / pretax_income で実際の税率を計算。異常値はフォールバック。
    eff_tax_rates = []
    for i in range(n):
        pt = g('pretax_income', i)
        tax = g('income_tax', i)
        if pt and pt != 0 and tax is not None:
            rate = abs(tax) / abs(pt)
            eff_tax_rates.append(rate if 0.01 < rate < 0.60 else 0.25)
        else:
            eff_tax_rates.append(0.25)
    eff_tax_now = eff_tax_rates[0] if eff_tax_rates else 0.25

    # ── 収益性指標（各年）────────────────────────────────────────────────────
    revenue = g_list('revenue')
    op_income = g_list('op_income')
    net_income = g_list('net_income')
    total_assets = g_list('total_assets')
    total_equity = g_list('total_equity')
    ebitda_list = g_list('ebitda')

    roe_list = []
    for i in range(n):
        ni = g('net_income', i)
        eq = g('total_equity', i)
        roe_list.append(ni / eq if (ni is not None and eq and eq != 0) else None)

    roa_list = []
    for i in range(n):
        ni = g('net_income', i)
        ta = g('total_assets', i)
        roa_list.append(ni / ta if (ni is not None and ta and ta != 0) else None)

    op_margin_list = []
    for i in range(n):
        oi = g('op_income', i)
        rev = g('revenue', i)
        op_margin_list.append(oi / rev if (oi is not None and rev and rev != 0) else None)

    ebitda_margin_list = []
    for i in range(n):
        eb = g('ebitda', i)
        rev = g('revenue', i)
        ebitda_margin_list.append(eb / rev if (eb is not None and rev and rev != 0) else None)

    def to_pct(v):
        return v * 100 if v is not None else None

    # ── 財務健全性指標 ────────────────────────────────────────────────────────
    equity_ratio = None
    equity_ratio_5y = None
    eq0 = g('total_equity', 0)
    ta0 = g('total_assets', 0)
    if eq0 and ta0 and ta0 != 0:
        equity_ratio = (eq0 / ta0) * 100
    i5 = min(4, n - 1) if n > 0 else 0
    eq4 = g('total_equity', i5)
    ta4 = g('total_assets', i5)
    if eq4 and ta4 and ta4 != 0:
        equity_ratio_5y = (eq4 / ta4) * 100

    current_r = None
    current_r_5y = None
    ca0 = g('current_assets', 0)
    cl0 = g('current_liab', 0)
    if ca0 and cl0 and cl0 != 0:
        current_r = (ca0 / cl0) * 100
    ca4 = g('current_assets', i5)
    cl4 = g('current_liab', i5)
    if ca4 and cl4 and cl4 != 0:
        current_r_5y = (ca4 / cl4) * 100

    quick_r = None
    quick_r_5y = None
    inv0 = g('inventory', 0)
    if ca0 and cl0 and cl0 != 0:
        quick_r = ((ca0 - (inv0 or 0)) / cl0) * 100
    inv4 = g('inventory', i5)
    if ca4 and cl4 and cl4 != 0:
        quick_r_5y = ((ca4 - (inv4 or 0)) / cl4) * 100

    # ── ROE/ROA サマリー ──────────────────────────────────────────────────────
    op_margin_vals = [to_pct(v) for v in op_margin_list[:5]]
    roe_now = to_pct(roe_list[0]) if roe_list else None
    roe_3y = to_pct(roe_list[2]) if len(roe_list) > 2 else None
    roe_5y = to_pct(roe_list[i5]) if roe_list and n > 0 else None
    roa_now = to_pct(roa_list[0]) if roa_list else None
    roa_3y = to_pct(roa_list[2]) if len(roa_list) > 2 else None
    roa_5y = to_pct(roa_list[i5]) if roa_list and n > 0 else None
    roe_growth = roe_now - roe_5y if (roe_now is not None and roe_5y is not None) else None

    ebitda_margin_val = to_pct(ebitda_margin_list[0]) if ebitda_margin_list else None
    ebitda_margin_5y = to_pct(ebitda_margin_list[i5]) if ebitda_margin_list and n > 0 else None

    # ── NOPAT（実効税率ベース）────────────────────────────────────────────────
    nopat = g('op_income', 0) * (1 - eff_tax_now) if g('op_income', 0) else None
    tax_rate_5y = eff_tax_rates[i5] if i5 < len(eff_tax_rates) else 0.25
    nopat_5y = g('op_income', i5) * (1 - tax_rate_5y) if n > 0 and g('op_income', i5) else None

    # ── 投下資本 ──────────────────────────────────────────────────────────────
    ic = g('invested_capital', 0)
    ic_5y = g('invested_capital', i5) if n > 0 else None
    if ic is None:
        eq = g('total_equity', 0)
        debt = g('total_debt', 0)
        cash = g('cash', 0)
        if eq is not None and debt is not None and cash is not None:
            ic = eq + debt - cash
    if ic_5y is None and n > 0:
        eq = g('total_equity', i5)
        debt = g('total_debt', i5)
        cash_v = g('cash', i5)
        if eq is not None and debt is not None and cash_v is not None:
            ic_5y = eq + debt - cash_v

    # ── WACC（CAPMベース）────────────────────────────────────────────────────
    # Ke = Rf + β × ERP
    # Kd = 支払利息 / 総負債（またはフォールバック3%）
    # WACC = E/V × Ke + D/V × Kd × (1 - 実効税率)
    beta = _safe(info.get('beta')) or 1.0

    # リスクフリーレート: 日本株=JGB~1%、米国株=^TNX（24hキャッシュ）
    rf_rate = _get_risk_free_rate(is_jpy)

    # 株式リスクプレミアム: 日本株~5%、米国株~5.5%（Damodaran推計）
    equity_premium = 0.050 if is_jpy else 0.055
    ke = rf_rate + beta * equity_premium

    wacc_val = None
    eq_val = g('total_equity', 0)
    debt_val = g('total_debt', 0)
    if eq_val and debt_val is not None:
        total_cap = eq_val + (debt_val or 0)
        if total_cap > 0:
            e_ratio = eq_val / total_cap
            d_ratio = (debt_val or 0) / total_cap
            # 実際の借入コスト: 支払利息 / 総負債
            kd = 0.03  # フォールバック
            int_exp = g('interest_exp', 0)
            if int_exp and debt_val and debt_val > 0:
                kd_calc = abs(int_exp) / debt_val
                if 0.001 < kd_calc < 0.20:
                    kd = kd_calc
            wacc_val = e_ratio * ke * 100 + d_ratio * kd * (1 - eff_tax_now) * 100
    elif eq_val:
        wacc_val = ke * 100  # 無借金企業

    # ── 販管費率 ──────────────────────────────────────────────────────────────
    sga_ratio = None
    sga_ratio_5y = None
    if g('sga', 0) and g('revenue', 0) and g('revenue', 0) != 0:
        sga_ratio = (g('sga', 0) / g('revenue', 0)) * 100
    if n > 0 and g('sga', i5) and g('revenue', i5):
        rev4 = g('revenue', i5)
        if rev4 and rev4 != 0:
            sga_ratio_5y = (g('sga', i5) / rev4) * 100

    # ── キャッシュフロー関連 ──────────────────────────────────────────────────
    fcf_list = g_list('fcf')
    ocf_list = g_list('ocf')
    capex_list = g_list('capex')

    debt_fcf = None
    debt_fcf_5y = None
    if g('total_debt', 0) is not None and g('fcf', 0) and g('fcf', 0) != 0:
        debt_fcf = g('total_debt', 0) / g('fcf', 0)
    if n > 0 and g('total_debt', i5) is not None and g('fcf', i5) and g('fcf', i5) != 0:
        debt_fcf_5y = g('total_debt', i5) / g('fcf', i5)

    nd_ebitda = None
    if g('net_debt', 0) is not None and g('ebitda', 0) and g('ebitda', 0) != 0:
        nd_ebitda = g('net_debt', 0) / g('ebitda', 0)

    debt_ebitda_val = None
    if g('total_debt', 0) is not None and g('ebitda', 0) and g('ebitda', 0) != 0:
        debt_ebitda_val = g('total_debt', 0) / g('ebitda', 0)

    # ── バリュエーション（info系）────────────────────────────────────────────
    per = _safe(info.get('trailingPE'))
    pbr = _safe(info.get('priceToBook'))
    div_yield = info.get('dividendYield')
    ev_val = info.get('enterpriseValue')
    company_name = info.get('shortName') or info.get('longName') or ticker_symbol
    industry = info.get('industry', '製造・サービス')
    dividend_yield_pct = div_yield * 100 if div_yield else None

    # 配当性向: info.get('payoutRatio') が小数 (0.30 = 30%)
    payout_ratio_now = _safe(info.get('payoutRatio'))
    if payout_ratio_now is not None:
        payout_ratio_now = payout_ratio_now * 100  # % 変換
    # 5年前の配当性向: 配当実績 / 純利益 から推計
    payout_ratio_5y = None
    if dividends_series is not None and len(dividends_series) > 0 and g('net_income', i5):
        try:
            div_5y_total = float(dividends_series[dividends_series.index.year == int(dates[i5])].sum()) if dates and i5 < len(dates) else 0
            shares_out = _safe(info.get('sharesOutstanding') or info.get('impliedSharesOutstanding'))
            if div_5y_total > 0 and shares_out and shares_out > 0:
                dps_5y = div_5y_total
                ni_5y = g('net_income', i5)
                eps_5y = g('eps', i5) or (g('eps_diluted', i5))
                if eps_5y and eps_5y != 0:
                    payout_ratio_5y = (dps_5y / eps_5y) * 100
        except Exception:
            pass

    # 5年前配当利回り: 配当履歴 / 株価履歴 から推計
    dividend_yield_5y = None
    if dividends_series is not None and len(dividends_series) > 0 and hist_df is not None and dates:
        try:
            year_5y_str = dates[i5] if i5 < len(dates) else None
            if year_5y_str:
                price_5y = _get_year_end_price(hist_df, year_5y_str, is_jpy)
                year_5y_int = int(year_5y_str)
                div_that_year = float(dividends_series[dividends_series.index.year == year_5y_int].sum())
                if price_5y and price_5y > 0 and div_that_year > 0:
                    dividend_yield_5y = (div_that_year / price_5y) * 100
        except Exception:
            pass

    # ── PER/PBR 5年履歴（株価履歴×EPS/BPSで計算）────────────────────────────
    shares_out = _safe(info.get('sharesOutstanding') or info.get('impliedSharesOutstanding'))

    per_ts = []
    pbr_ts = []
    for i in range(min(5, n)):
        price_y = _get_year_end_price(hist_df, dates[i], is_jpy) if dates and i < len(dates) else None
        # PER = 株価 / EPS
        eps_y = g('eps', i) or g('eps_diluted', i)
        if price_y and eps_y and eps_y > 0:
            per_ts.append(_safe(price_y / eps_y))
        else:
            per_ts.append(None)
        # PBR = 株価 / BPS (BPS = 純資産 / 発行済株式数)
        eq_y = g('total_equity', i)
        if price_y and eq_y and shares_out and shares_out > 0:
            bps = eq_y / shares_out
            if bps > 0:
                pbr_ts.append(_safe(price_y / bps))
            else:
                pbr_ts.append(None)
        else:
            pbr_ts.append(None)

    per_5y_val = per_ts[i5] if i5 < len(per_ts) else None
    pbr_5y_val = pbr_ts[i5] if i5 < len(pbr_ts) else None
    # 現在のPER/PBRはinfoの値を優先、なければ計算値
    per_now = per if per is not None else (per_ts[0] if per_ts else None)
    pbr_now = pbr if pbr is not None else (pbr_ts[0] if pbr_ts else None)

    # ── アナリスト評価 ────────────────────────────────────────────────────────
    analyst_consensus, analyst_breakdown = _get_analyst_consensus(recommendations_df)
    analyst_target = _safe(info.get('targetMeanPrice'))
    current_price = _safe(info.get('currentPrice') or info.get('regularMarketPrice'))
    analyst_upside = None
    if analyst_target and current_price and current_price > 0:
        analyst_upside = ((analyst_target - current_price) / current_price) * 100

    # ── 配当成長率（CAGR）────────────────────────────────────────────────────
    div_growth_rate = _calc_dividend_growth_rate(dividends_series)

    # ── ESGリスク評価 ─────────────────────────────────────────────────────────
    d3_esg = _assess_esg(sustainability_df)

    # ── 株主構造評価 ──────────────────────────────────────────────────────────
    d2_ownership = _assess_ownership(major_holders_df)

    # ── data dict (parse_excel互換) ───────────────────────────────────────────
    data = {
        "company": company_name,
        "ticker": ticker_symbol,
        "industry": industry,

        "revenue": [g('revenue', i) for i in range(min(5, n))],
        "fcf": [g('fcf', i) for i in range(min(5, len(fcf_list)))],
        "eps": [g('eps', i) or g('eps_diluted', i) for i in range(min(5, n))],

        "roe": [roe_now, roe_3y, roe_5y],
        "roe_growth_rate": roe_growth,
        "roa": [roa_now, roa_3y, roa_5y],

        "equity_ratio": equity_ratio,
        "equity_ratio_5y": equity_ratio_5y,
        "quick_ratio": quick_r,
        "quick_ratio_5y": quick_r_5y,
        "current_ratio": current_r,
        "current_ratio_5y": current_r_5y,

        "operating_cf": [g('ocf', i) for i in range(min(5, len(ocf_list)))],
        "investing_cf": [g('investing_cf', i) for i in range(min(5, len(g_list('investing_cf'))))],
        "financing_cf": [g('financing_cf', i) for i in range(min(5, len(g_list('financing_cf'))))],
        "op_margin": op_margin_vals,
        "ebitda_margin": ebitda_margin_val,
        "ebitda_margin_5y": ebitda_margin_5y,

        "debt_fcf": debt_fcf,
        "debt_fcf_5y": debt_fcf_5y,
        "nd_ebitda": nd_ebitda,
        "ev": _safe(ev_val),
        "per": per_now,
        "per_5y": per_5y_val,
        "pbr": pbr_now,
        "pbr_5y": pbr_5y_val,

        "nopat": nopat,
        "nopat_5y": nopat_5y,
        "invested_capital": ic,
        "invested_capital_5y": ic_5y,
        "wacc": wacc_val,

        "accounts_receivable": g('receivables', 0),
        "accounts_receivable_5y": g('receivables', i5) if n > 0 else None,
        "inventory": g('inventory', 0),
        "inventory_5y": g('inventory', i5) if n > 0 else None,
        "accounts_payable": g('payables', 0),
        "accounts_payable_5y": g('payables', i5) if n > 0 else None,
        "cogs": g('cogs', 0),
        "cogs_5y": g('cogs', i5) if n > 0 else None,
        "sga_ratio": sga_ratio,
        "sga_ratio_5y": sga_ratio_5y,

        "total_assets": g('total_assets', 0),
        "total_assets_5y": g('total_assets', i5) if n > 0 else None,
        "fixed_assets": g('fixed_assets', 0),
        "fixed_assets_5y": g('fixed_assets', i5) if n > 0 else None,
        "tangible_fixed_assets": g('fixed_assets', 0),
        "tangible_fixed_assets_5y": g('fixed_assets', i5) if n > 0 else None,
        "intangible_fixed_assets": g('intangibles', 0),
        "intangible_fixed_assets_5y": g('intangibles', i5) if n > 0 else None,

        "net_income_val": g('net_income', 0),
        "net_income_val_5y": g('net_income', i5) if n > 0 else None,
        "op_income_val": g('op_income', 0),
        "op_income_val_5y": g('op_income', i5) if n > 0 else None,
        "interest_exp": g('interest_exp', 0),
        "interest_exp_5y": g('interest_exp', i5) if n > 0 else None,
        "other_exp": g('other_exp', 0),
        "other_exp_5y": g('other_exp', i5) if n > 0 else None,
        "pretax_income": g('pretax_income', 0),
        "pretax_income_5y": g('pretax_income', i5) if n > 0 else None,
        "income_tax": g('income_tax', 0),
        "income_tax_5y": g('income_tax', i5) if n > 0 else None,
        "total_equity": g('total_equity', 0),
        "total_equity_5y": g('total_equity', i5) if n > 0 else None,

        "dividend_yield": dividend_yield_pct,
        "dividend_yield_5y": dividend_yield_5y,
        "payout_ratio": payout_ratio_now,
        "payout_ratio_5y": payout_ratio_5y,
        "dividend_growth_rate": div_growth_rate,

        "d1_mgmt_change": "○",
        "d2_ownership": d2_ownership,
        "d3_esg": d3_esg,
    }

    # ── ts_data dict (時系列、parse_excel互換) ────────────────────────────────
    da_list = g_list('da')
    sga_list_ts = g_list('sga')
    investing_cf_list = g_list('investing_cf')
    financing_cf_list = g_list('financing_cf')

    current_ratio_ts = []
    for i in range(n):
        ca = g('current_assets', i)
        cl = g('current_liab', i)
        current_ratio_ts.append((ca / cl) * 100 if ca and cl and cl != 0 else None)

    quick_ratio_ts = []
    for i in range(n):
        ca = g('current_assets', i)
        cl = g('current_liab', i)
        inv = g('inventory', i) or 0
        quick_ratio_ts.append(((ca - inv) / cl) * 100 if ca and cl and cl != 0 else None)

    equity_ratio_ts = []
    for i in range(n):
        eq = g('total_equity', i)
        ta = g('total_assets', i)
        equity_ratio_ts.append((eq / ta) * 100 if eq and ta and ta != 0 else None)

    # ROIC（実効税率ベース）
    roic_ts = []
    for i in range(n):
        oi = g('op_income', i)
        ic_i = g('invested_capital', i)
        if ic_i is None:
            eq = g('total_equity', i)
            debt = g('total_debt', i)
            cash_i = g('cash', i)
            if eq is not None and debt is not None and cash_i is not None:
                ic_i = eq + debt - cash_i
        tax_i = eff_tax_rates[i] if i < len(eff_tax_rates) else 0.25
        if oi is not None and ic_i and ic_i != 0:
            roic_ts.append((oi * (1 - tax_i) / ic_i) * 100)
        else:
            roic_ts.append(None)

    debt_fcf_ts = []
    for i in range(n):
        debt = g('total_debt', i)
        fcf_i = g('fcf', i)
        debt_fcf_ts.append(debt / fcf_i if (debt is not None and fcf_i and fcf_i != 0) else None)

    debt_ebitda_ts = []
    for i in range(n):
        debt = g('total_debt', i)
        eb = g('ebitda', i)
        debt_ebitda_ts.append(debt / eb if (debt is not None and eb and eb != 0) else None)

    nd_ebitda_ts = []
    for i in range(n):
        nd = g('net_debt', i)
        eb = g('ebitda', i)
        nd_ebitda_ts.append(nd / eb if (nd is not None and eb and eb != 0) else None)

    ts_data = {
        "dates": dates,
        "revenue": list(revenue),
        "net_income": list(net_income),
        "fcf": list(fcf_list),
        "eps": [g('eps', i) or g('eps_diluted', i) for i in range(n)],
        "ocf": list(ocf_list),
        "investing_cf": list(investing_cf_list),
        "financing_cf": list(financing_cf_list),
        "ebitda": list(ebitda_list),
        "total_assets": list(total_assets),
        "total_equity": list(total_equity),
        "total_debt": list(g_list('total_debt')),
        "roe": [to_pct(v) for v in roe_list],
        "roa": [to_pct(v) for v in roa_list],
        "op_margin": [to_pct(v) for v in op_margin_list],
        "quick_ratio": quick_ratio_ts,
        "current_ratio": current_ratio_ts,
        "equity_ratio": equity_ratio_ts,
        "ebitda_margin": [to_pct(v) for v in ebitda_margin_list],
        "debt_fcf": debt_fcf_ts,
        "roic": roic_ts,
        "capex": list(capex_list),
        "sga": list(sga_list_ts),
        "da": list(da_list),
        "pe_ratio": per_ts,
        "pb_ratio": pbr_ts,
        "debt_ebitda": debt_ebitda_ts,
        "nd_ebitda": nd_ebitda_ts,
        "dividend_yield": [],
        "payout_ratio": [],
        "eff_tax_rate": eff_tax_rates,
    }

    # ── DuPont分解 ────────────────────────────────────────────────────────────
    net_margin_ts = []
    asset_turnover_ts = []
    fin_leverage_ts = []
    for i in range(n):
        ni = g('net_income', i)
        rev = g('revenue', i)
        ta = g('total_assets', i)
        eq = g('total_equity', i)
        net_margin_ts.append(round(ni / rev * 100, 2) if ni is not None and rev and rev != 0 else None)
        asset_turnover_ts.append(round(rev / ta, 3) if rev is not None and ta and ta != 0 else None)
        fin_leverage_ts.append(round(ta / eq, 3) if ta is not None and eq and eq != 0 else None)

    ts_data["net_margin"] = net_margin_ts
    ts_data["asset_turnover"] = asset_turnover_ts
    ts_data["financial_leverage"] = fin_leverage_ts

    # ── 純利益率分解 ──────────────────────────────────────────────────────────
    interest_burden_ts = []
    tax_burden_ts = []
    nonop_burden_ts = []
    for i in range(n):
        oi = g('op_income', i)
        pt = g('pretax_income', i)
        ni = g('net_income', i)
        ie = g('interest_exp', i)

        if oi is not None and oi != 0 and ie is not None:
            interest_burden_ts.append(round((oi + ie) / oi * 100, 2))
        else:
            interest_burden_ts.append(None)

        if oi is not None and ie is not None and (oi + ie) != 0 and pt is not None:
            nonop_burden_ts.append(round(pt / (oi + ie) * 100, 2))
        else:
            nonop_burden_ts.append(None)

        if pt is not None and pt != 0 and ni is not None:
            tax_burden_ts.append(round(ni / pt * 100, 2))
        else:
            tax_burden_ts.append(None)

    ts_data["interest_burden"] = interest_burden_ts
    ts_data["nonop_burden"] = nonop_burden_ts
    ts_data["tax_burden"] = tax_burden_ts

    # ── アナリスト・メタデータ ────────────────────────────────────────────────
    ts_data["_source"] = "yfinance"
    ts_data["_currency"] = currency
    ts_data["_country"] = info.get("country", "US")
    ts_data["_is_jpy"] = is_jpy
    ts_data["_beta"] = beta
    ts_data["_risk_free_rate"] = rf_rate * 100
    ts_data["_equity_premium"] = equity_premium * 100
    ts_data["_analyst_consensus"] = analyst_consensus
    ts_data["_analyst_breakdown"] = analyst_breakdown
    ts_data["_analyst_target"] = analyst_target
    ts_data["_analyst_upside"] = analyst_upside
    ts_data["_current_price"] = current_price
    ts_data["_eff_tax_rate_now"] = round(eff_tax_now * 100, 1)

    return data, ts_data
