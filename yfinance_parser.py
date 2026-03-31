"""
yfinanceからティッカーシンボルで財務データを取得し、
parse_excel()と同一形式の (data, ts_data) を返すパーサー
"""
import math
import yfinance as yf


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

    # データ取得
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

    # 各財務諸表からデータ抽出
    inc_data, inc_dates = _extract_series(inc_df, _INCOME_MAP)
    cf_data, cf_dates = _extract_series(cf_df, _CASHFLOW_MAP)
    bs_data, bs_dates = _extract_series(bs_df, _BALANCE_MAP)

    # 統合辞書（incの日付を基準）
    dates = inc_dates or cf_dates or bs_dates
    all_data = {}
    all_data.update(bs_data)
    all_data.update(cf_data)
    all_data.update(inc_data)  # incが最優先

    # ヘルパー
    def g(key, idx=0):
        lst = all_data.get(key, [])
        return lst[idx] if idx < len(lst) else None

    def g_list(key):
        return all_data.get(key, [])

    n = len(dates)

    # --- 計算指標 ---

    revenue = g_list('revenue')
    op_income = g_list('op_income')
    net_income = g_list('net_income')
    total_assets = g_list('total_assets')
    total_equity = g_list('total_equity')
    ebitda_list = g_list('ebitda')

    # ROE (各年)
    roe_list = []
    for i in range(n):
        ni = g('net_income', i)
        eq = g('total_equity', i)
        if ni is not None and eq and eq != 0:
            roe_list.append(ni / eq)
        else:
            roe_list.append(None)

    # ROA (各年)
    roa_list = []
    for i in range(n):
        ni = g('net_income', i)
        ta = g('total_assets', i)
        if ni is not None and ta and ta != 0:
            roa_list.append(ni / ta)
        else:
            roa_list.append(None)

    # Operating Margin (各年)
    op_margin_list = []
    for i in range(n):
        oi = g('op_income', i)
        rev = g('revenue', i)
        if oi is not None and rev and rev != 0:
            op_margin_list.append(oi / rev)
        else:
            op_margin_list.append(None)

    # EBITDA Margin (各年)
    ebitda_margin_list = []
    for i in range(n):
        eb = g('ebitda', i)
        rev = g('revenue', i)
        if eb is not None and rev and rev != 0:
            ebitda_margin_list.append(eb / rev)
        else:
            ebitda_margin_list.append(None)

    # %変換ヘルパー
    def to_pct(v):
        return v * 100 if v is not None else None

    # Equity Ratio
    equity_ratio = None
    equity_ratio_5y = None
    eq0 = g('total_equity', 0)
    ta0 = g('total_assets', 0)
    if eq0 and ta0 and ta0 != 0:
        equity_ratio = (eq0 / ta0) * 100
    eq4 = g('total_equity', min(4, n - 1)) if n > 0 else None
    ta4 = g('total_assets', min(4, n - 1)) if n > 0 else None
    if eq4 and ta4 and ta4 != 0:
        equity_ratio_5y = (eq4 / ta4) * 100

    # Current Ratio / Quick Ratio
    current_r = None
    current_r_5y = None
    ca0 = g('current_assets', 0)
    cl0 = g('current_liab', 0)
    if ca0 and cl0 and cl0 != 0:
        current_r = (ca0 / cl0) * 100
    ca4 = g('current_assets', min(4, n - 1)) if n > 0 else None
    cl4 = g('current_liab', min(4, n - 1)) if n > 0 else None
    if ca4 and cl4 and cl4 != 0:
        current_r_5y = (ca4 / cl4) * 100

    # Quick Ratio = (Current Assets - Inventory) / Current Liab
    quick_r = None
    quick_r_5y = None
    inv0 = g('inventory', 0)
    if ca0 and cl0 and cl0 != 0:
        quick_r = ((ca0 - (inv0 or 0)) / cl0) * 100
    inv4 = g('inventory', min(4, n - 1)) if n > 0 else None
    if ca4 and cl4 and cl4 != 0:
        quick_r_5y = ((ca4 - (inv4 or 0)) / cl4) * 100

    # Op Margin %
    op_margin_vals = [to_pct(v) for v in op_margin_list[:5]]

    # ROE/ROA
    roe_now = to_pct(g_list('_roe')[0]) if '_roe' in all_data else to_pct(roe_list[0] if roe_list else None)
    roe_now = to_pct(roe_list[0]) if roe_list and len(roe_list) > 0 else None
    roe_3y = to_pct(roe_list[2]) if len(roe_list) > 2 else None
    roe_5y = to_pct(roe_list[min(4, n - 1)]) if roe_list and n > 0 else None
    roa_now = to_pct(roa_list[0]) if roa_list and len(roa_list) > 0 else None
    roa_3y = to_pct(roa_list[2]) if len(roa_list) > 2 else None
    roa_5y = to_pct(roa_list[min(4, n - 1)]) if roa_list and n > 0 else None

    roe_growth = roe_now - roe_5y if (roe_now is not None and roe_5y is not None) else None

    # EBITDA Margin
    ebitda_margin_val = to_pct(ebitda_margin_list[0]) if ebitda_margin_list else None
    ebitda_margin_5y = to_pct(ebitda_margin_list[min(4, n - 1)]) if ebitda_margin_list and n > 0 else None

    # NOPAT
    nopat = g('op_income', 0) * 0.75 if g('op_income', 0) else None
    nopat_5y = g('op_income', min(4, n - 1)) * 0.75 if n > 0 and g('op_income', min(4, n - 1)) else None

    # Invested Capital
    ic = g('invested_capital', 0)
    ic_5y = g('invested_capital', min(4, n - 1)) if n > 0 else None
    # フォールバック: Equity + Debt - Cash
    if ic is None:
        eq = g('total_equity', 0)
        debt = g('total_debt', 0)
        cash = g('cash', 0)
        if eq is not None and debt is not None and cash is not None:
            ic = eq + debt - cash
    if ic_5y is None and n > 0:
        idx = min(4, n - 1)
        eq = g('total_equity', idx)
        debt = g('total_debt', idx)
        cash = g('cash', idx)
        if eq is not None and debt is not None and cash is not None:
            ic_5y = eq + debt - cash

    # WACC
    wacc_val = None
    if g('total_equity', 0) and g('total_debt', 0):
        total_cap = g('total_equity', 0) + g('total_debt', 0)
        if total_cap > 0:
            d_ratio = g('total_debt', 0) / total_cap
            e_ratio = 1 - d_ratio
            wacc_val = e_ratio * 8.0 + d_ratio * 3.0 * 0.75

    # SGA Ratio
    sga_ratio = None
    sga_ratio_5y = None
    if g('sga', 0) and g('revenue', 0) and g('revenue', 0) != 0:
        sga_ratio = (g('sga', 0) / g('revenue', 0)) * 100
    if n > 0 and g('sga', min(4, n - 1)) and g('revenue', min(4, n - 1)):
        rev4 = g('revenue', min(4, n - 1))
        if rev4 and rev4 != 0:
            sga_ratio_5y = (g('sga', min(4, n - 1)) / rev4) * 100

    # Debt/Equity for equity_ratio fallback
    debt_equity_list = []
    for i in range(n):
        debt = g('total_debt', i)
        eq = g('total_equity', i)
        if debt is not None and eq and eq != 0:
            debt_equity_list.append(debt / eq)
        else:
            debt_equity_list.append(None)

    # info系指標
    per = info.get('trailingPE')
    pbr = info.get('priceToBook')
    div_yield = info.get('dividendYield')
    ev_val = info.get('enterpriseValue')
    company_name = info.get('shortName') or info.get('longName') or ticker_symbol
    industry = info.get('industry', '製造・サービス')

    # dividendYieldはyfinanceでは小数 (0.0279 = 2.79%)
    dividend_yield_pct = div_yield * 100 if div_yield else None

    # 5年前インデックス
    i5 = min(4, n - 1) if n > 0 else 0

    # FCF
    fcf_list = g_list('fcf')
    ocf_list = g_list('ocf')
    capex_list = g_list('capex')

    # Debt/FCF
    debt_fcf = None
    debt_fcf_5y = None
    if g('total_debt', 0) is not None and g('fcf', 0) and g('fcf', 0) != 0:
        debt_fcf = g('total_debt', 0) / g('fcf', 0)
    if n > 0 and g('total_debt', i5) is not None and g('fcf', i5) and g('fcf', i5) != 0:
        debt_fcf_5y = g('total_debt', i5) / g('fcf', i5)

    # Net Debt / EBITDA
    nd_ebitda = None
    if g('net_debt', 0) is not None and g('ebitda', 0) and g('ebitda', 0) != 0:
        nd_ebitda = g('net_debt', 0) / g('ebitda', 0)

    # Debt/EBITDA
    debt_ebitda_val = None
    if g('total_debt', 0) is not None and g('ebitda', 0) and g('ebitda', 0) != 0:
        debt_ebitda_val = g('total_debt', 0) / g('ebitda', 0)

    # --- data dict (parse_excel互換) ---
    data = {
        "company": company_name,
        "ticker": ticker_symbol,
        "industry": industry,

        "revenue": [g('revenue', i) for i in range(min(5, n))],
        "fcf": [g('fcf', i) for i in range(min(5, len(fcf_list)))],
        "eps": [g('eps', i) for i in range(min(5, len(g_list('eps'))))],

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
        "per": _safe(per),
        "per_5y": None,
        "pbr": _safe(pbr),
        "pbr_5y": None,

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
        "dividend_yield_5y": None,
        "payout_ratio": None,
        "payout_ratio_5y": None,

        "d1_mgmt_change": "○",
        "d2_ownership": "○",
        "d3_esg": "○",
    }

    # --- ts_data dict (時系列、parse_excel互換) ---
    da_list = g_list('da')
    sga_list_ts = g_list('sga')
    investing_cf_list = g_list('investing_cf')
    financing_cf_list = g_list('financing_cf')

    # Current Ratio 時系列
    current_ratio_ts = []
    for i in range(n):
        ca = g('current_assets', i)
        cl = g('current_liab', i)
        if ca and cl and cl != 0:
            current_ratio_ts.append((ca / cl) * 100)
        else:
            current_ratio_ts.append(None)

    # Quick Ratio 時系列
    quick_ratio_ts = []
    for i in range(n):
        ca = g('current_assets', i)
        cl = g('current_liab', i)
        inv = g('inventory', i) or 0
        if ca and cl and cl != 0:
            quick_ratio_ts.append(((ca - inv) / cl) * 100)
        else:
            quick_ratio_ts.append(None)

    # Equity Ratio 時系列
    equity_ratio_ts = []
    for i in range(n):
        eq = g('total_equity', i)
        ta = g('total_assets', i)
        if eq and ta and ta != 0:
            equity_ratio_ts.append((eq / ta) * 100)
        else:
            equity_ratio_ts.append(None)

    # ROIC 時系列
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
        if oi is not None and ic_i and ic_i != 0:
            roic_ts.append((oi * 0.75 / ic_i) * 100)
        else:
            roic_ts.append(None)

    # Debt/FCF 時系列
    debt_fcf_ts = []
    for i in range(n):
        debt = g('total_debt', i)
        fcf_i = g('fcf', i)
        if debt is not None and fcf_i and fcf_i != 0:
            debt_fcf_ts.append(debt / fcf_i)
        else:
            debt_fcf_ts.append(None)

    # Debt/EBITDA 時系列
    debt_ebitda_ts = []
    for i in range(n):
        debt = g('total_debt', i)
        eb = g('ebitda', i)
        if debt is not None and eb and eb != 0:
            debt_ebitda_ts.append(debt / eb)
        else:
            debt_ebitda_ts.append(None)

    # Net Debt/EBITDA 時系列
    nd_ebitda_ts = []
    for i in range(n):
        nd = g('net_debt', i)
        eb = g('ebitda', i)
        if nd is not None and eb and eb != 0:
            nd_ebitda_ts.append(nd / eb)
        else:
            nd_ebitda_ts.append(None)

    ts_data = {
        "dates": dates,
        "revenue": list(revenue),
        "net_income": list(net_income),
        "fcf": list(fcf_list),
        "eps": list(g_list('eps')),
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
        "pe_ratio": [],
        "pb_ratio": [],
        "debt_ebitda": debt_ebitda_ts,
        "nd_ebitda": nd_ebitda_ts,
        "dividend_yield": [],
        "payout_ratio": [],
    }

    # DuPont分解
    net_margin_ts = []
    asset_turnover_ts = []
    fin_leverage_ts = []
    for i in range(n):
        ni = g('net_income', i)
        rev = g('revenue', i)
        ta = g('total_assets', i)
        eq = g('total_equity', i)
        if ni is not None and rev and rev != 0:
            net_margin_ts.append(round(ni / rev * 100, 2))
        else:
            net_margin_ts.append(None)
        if rev is not None and ta and ta != 0:
            asset_turnover_ts.append(round(rev / ta, 3))
        else:
            asset_turnover_ts.append(None)
        if ta is not None and eq and eq != 0:
            fin_leverage_ts.append(round(ta / eq, 3))
        else:
            fin_leverage_ts.append(None)

    ts_data["net_margin"] = net_margin_ts
    ts_data["asset_turnover"] = asset_turnover_ts
    ts_data["financial_leverage"] = fin_leverage_ts

    # 純利益率分解
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

    # メタデータ（フロントでのフォーマット判定用）
    ts_data["_source"] = "yfinance"
    currency = info.get("currency", "USD")
    ts_data["_currency"] = currency
    ts_data["_country"] = info.get("country", "US")
    # 通貨から地域を自動判定
    ts_data["_is_jpy"] = currency == "JPY" or ticker_symbol.endswith(".T") or ticker_symbol.endswith(".J")

    return data, ts_data
