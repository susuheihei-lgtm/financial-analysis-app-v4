"""
Excelファイルからstock_data.json形式のデータを抽出するパーサー
シートが不足していても利用可能なデータだけで分析を行う
"""
import openpyxl


def _get_row_data(ws, row_label):
    """指定ラベルの行データを取得（新しい順）"""
    if ws is None:
        return []
    for r in range(1, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == row_label:
            vals = []
            for c in range(2, ws.max_column + 1):
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    vals.append(v)
            return list(reversed(vals))  # 新しい順
    return []


def _safe_get(lst, idx, default=None):
    return lst[idx] if idx < len(lst) else default


def _find_sheet(wb, candidates):
    """候補名リストからシートを探す。見つからなければNone"""
    for name in candidates:
        if name in wb.sheetnames:
            return wb[name]
    # 部分一致でも探す
    lower_sheets = {s.lower(): s for s in wb.sheetnames}
    for name in candidates:
        for key, real_name in lower_sheets.items():
            if name.lower() in key:
                return wb[real_name]
    return None


def _try_labels(ws, labels):
    """複数のラベル候補から最初にヒットしたデータを返す"""
    for label in labels:
        data = _get_row_data(ws, label)
        if data:
            return data
    return []


def parse_excel(filepath):
    """Excelファイルをパースしてanalyzer用のdictを返す。
    シートが存在しない場合はそのシートのデータを空として扱う。"""
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # シートの自動検出（複数の名前パターンに対応）
    inc = _find_sheet(wb, ['Income-Annual', 'Income Statement', 'Income', 'Export', 'income'])
    bs = _find_sheet(wb, ['Balance-Sheet-Annual', 'Balance Sheet', 'Balance', 'balance'])
    cf = _find_sheet(wb, ['Cash-Flow-Annual', 'Cash Flow', 'Cash Flow Statement', 'cashflow'])
    rat = _find_sheet(wb, ['Ratios-Annual', 'Ratios', 'Financial Ratios', 'ratios'])

    # もし全シートNoneなら、最初のシートをincとして使う
    if inc is None and bs is None and cf is None and rat is None:
        if wb.sheetnames:
            inc = wb[wb.sheetnames[0]]

    # 日付（新しい順）- 複数のラベルに対応
    dates_raw = _try_labels(inc, ['Date', 'Year Ending', 'Fiscal Year', 'Period'])

    # 収益データ
    revenue = _try_labels(inc, ['Revenue', 'Total Revenue', 'Net Revenue', 'Sales'])
    cogs_list = _try_labels(inc, ['Cost of Revenue', 'Cost of Goods Sold', 'COGS'])
    op_income = _try_labels(inc, ['Operating Income', 'Operating Profit', 'EBIT'])
    net_income = _try_labels(inc, ['Net Income', 'Net Profit', 'Net Earnings'])
    eps_list = _try_labels(inc, ['EPS (Basic)', 'EPS', 'Earnings Per Share'])
    op_margin_list = _try_labels(inc, ['Operating Margin'])
    ebitda_margin_list = _try_labels(inc, ['EBITDA Margin'])
    ebitda_list = _try_labels(inc, ['EBITDA'])
    sga_list = _try_labels(inc, ['Selling, General & Admin', 'SG&A', 'SGA'])
    da_list = _try_labels(inc, ['Depreciation & Amortization', 'D&A', 'Depreciation'])
    interest_exp_list = _try_labels(inc, ['Interest Expense / Income', 'Interest Expense', 'Net Interest Income'])
    other_exp_list = _try_labels(inc, ['Other Expense / Income', 'Other Income/Expense', 'Non-Operating Income'])
    pretax_income_list = _try_labels(inc, ['Pretax Income', 'Pre-Tax Income', 'Income Before Tax'])
    income_tax_list = _try_labels(inc, ['Income Tax', 'Tax Provision', 'Provision for Income Taxes'])
    eff_tax_rate_list = _try_labels(inc, ['Effective Tax Rate'])

    # キャッシュフロー
    fcf_list = _try_labels(cf, ['Free Cash Flow', 'FCF'])
    ocf_list = _try_labels(cf, ['Operating Cash Flow', 'Cash from Operations'])
    capex_list = _try_labels(cf, ['Capital Expenditures', 'Capex', 'CapEx'])
    investing_cf_list = _try_labels(cf, ['Investing Cash Flow', 'Cash from Investing'])
    financing_cf_list = _try_labels(cf, ['Financing Cash Flow', 'Cash from Financing'])

    # バランスシート
    total_assets_list = _try_labels(bs, ['Total Assets'])
    total_equity_list = _try_labels(bs, ['Shareholders Equity', "Shareholders' Equity", 'Total Equity', 'Stockholders Equity'])
    total_debt_list = _try_labels(bs, ['Total Debt', 'Long-Term Debt'])
    receivables_list = _try_labels(bs, ['Receivables', 'Accounts Receivable', 'Trade Receivables'])
    inventory_list = _try_labels(bs, ['Inventory', 'Inventories'])
    payables_list = _try_labels(bs, ['Accounts Payable', 'Trade Payables'])
    current_assets_list = _try_labels(bs, ['Total Current Assets', 'Current Assets'])
    current_liab_list = _try_labels(bs, ['Total Current Liabilities', 'Current Liabilities'])
    cash_list = _try_labels(bs, ['Cash & Cash Equivalents', 'Cash and Equivalents', 'Cash'])
    fixed_assets_list = _try_labels(bs, ['Property, Plant & Equipment', 'PP&E', 'Fixed Assets'])
    intangibles_list = _try_labels(bs, ['Goodwill and Intangibles', 'Intangible Assets', 'Goodwill'])
    net_debt_list = _try_labels(bs, ['Net Cash (Debt)', 'Net Debt'])
    long_term_assets_list = _try_labels(bs, ['Total Long-Term Assets', 'Non-Current Assets'])

    # 比率
    pe_list = _try_labels(rat, ['PE Ratio', 'P/E Ratio', 'Price/Earnings'])
    pb_list = _try_labels(rat, ['PB Ratio', 'P/B Ratio', 'Price/Book'])
    ev_list = _try_labels(rat, ['Enterprise Value', 'EV'])
    roe_list = _try_labels(rat, ['Return on Equity (ROE)', 'ROE', 'Return on Equity'])
    roa_list = _try_labels(rat, ['Return on Assets (ROA)', 'ROA', 'Return on Assets'])
    current_ratio_list = _try_labels(rat, ['Current Ratio'])
    quick_ratio_list = _try_labels(rat, ['Quick Ratio'])
    debt_fcf_list = _try_labels(rat, ['Debt/FCF'])
    debt_ebitda_list = _try_labels(rat, ['Debt/EBITDA'])
    nd_ebitda_list = _try_labels(rat, ['Net Debt/EBITDA'])
    roic_list = _try_labels(rat, ['Return on Invested Capital (ROIC)', 'ROIC'])
    debt_equity_list = _try_labels(rat, ['Debt/Equity', 'D/E Ratio'])
    dividend_yield_list = _try_labels(rat, ['Dividend Yield'])
    payout_ratio_list = _try_labels(rat, ['Payout Ratio'])

    # ROE/ROAがRatiosシートにないがIncome+BSから計算可能な場合
    if not roe_list and net_income and total_equity_list:
        roe_list = []
        max_len_roe = min(len(net_income), len(total_equity_list))
        for i in range(max_len_roe):
            ni = _safe_get(net_income, i)
            eq = _safe_get(total_equity_list, i)
            if ni is not None and eq and eq != 0:
                roe_list.append(ni / eq)
            else:
                roe_list.append(None)

    if not roa_list and net_income and total_assets_list:
        roa_list = []
        max_len_roa = min(len(net_income), len(total_assets_list))
        for i in range(max_len_roa):
            ni = _safe_get(net_income, i)
            ta = _safe_get(total_assets_list, i)
            if ni is not None and ta and ta != 0:
                roa_list.append(ni / ta)
            else:
                roa_list.append(None)

    # 値の取得ヘルパー
    def g(lst, idx):
        return _safe_get(lst, idx)

    # パーセント変換
    def to_pct(v):
        return v * 100 if v is not None else None

    # 自己資本比率の計算 D/Eベース: 1/(1+D/E)*100  (業種を問わず比較可能)
    equity_ratio = None
    equity_ratio_5y = None
    de0 = g(debt_equity_list, 0)
    de4 = g(debt_equity_list, 4)
    if de0 is not None:
        equity_ratio = (1 / (1 + de0)) * 100
    elif g(total_equity_list, 0) and g(total_assets_list, 0):
        equity_ratio = (g(total_equity_list, 0) / g(total_assets_list, 0)) * 100
    if de4 is not None:
        equity_ratio_5y = (1 / (1 + de4)) * 100
    elif g(total_equity_list, 4) and g(total_assets_list, 4):
        equity_ratio_5y = (g(total_equity_list, 4) / g(total_assets_list, 4)) * 100

    # 当座比率 (現在値 + 5年前)
    quick_r = to_pct(g(quick_ratio_list, 0))
    quick_r_5y = to_pct(g(quick_ratio_list, 4))

    # 流動比率 (現在値 + 5年前)
    current_r = to_pct(g(current_ratio_list, 0))
    current_r_5y = to_pct(g(current_ratio_list, 4))

    # 営業利益率 （%単位に変換）— 5年分
    op_margin_vals = [to_pct(g(op_margin_list, i)) for i in range(min(5, len(op_margin_list)))]
    # 営業利益率がRatiosにない場合、計算で補完
    if not op_margin_vals and op_income and revenue:
        op_margin_vals = []
        for i in range(min(5, len(op_income))):
            oi = g(op_income, i)
            rev = g(revenue, i)
            if oi is not None and rev and rev != 0:
                op_margin_vals.append(round(oi / rev * 100, 2))
            else:
                op_margin_vals.append(None)

    # EBITDAマージン (5年前)
    ebitda_margin_5y = to_pct(g(ebitda_margin_list, 4))

    # EBITDAマージン
    ebitda_margin_val = to_pct(g(ebitda_margin_list, 0))

    # ROE, ROA（%単位に変換）
    roe_now = to_pct(g(roe_list, 0))
    roe_3y = to_pct(g(roe_list, 2))
    roe_5y = to_pct(g(roe_list, 4))
    roa_now = to_pct(g(roa_list, 0))
    roa_3y = to_pct(g(roa_list, 2))
    roa_5y = to_pct(g(roa_list, 4))

    # ROE成長率
    roe_growth = roe_now - roe_5y if (roe_now is not None and roe_5y is not None) else None

    # NOPAT = Operating Income * (1 - Tax Rate)
    nopat = g(op_income, 0) * 0.75 if g(op_income, 0) else None
    nopat_5y = g(op_income, 4) * 0.75 if g(op_income, 4) else None

    # 投下資本 = Total Equity + Total Debt - Cash
    def calc_ic(eq, debt, cash):
        if eq is not None and debt is not None and cash is not None:
            return eq + debt - cash
        return None

    ic = calc_ic(g(total_equity_list, 0), g(total_debt_list, 0), g(cash_list, 0))
    ic_5y = calc_ic(g(total_equity_list, 4), g(total_debt_list, 4), g(cash_list, 4))

    # WACC簡易推定
    wacc_val = None
    if g(total_equity_list, 0) and g(total_debt_list, 0) and g(total_assets_list, 0):
        d_ratio = g(total_debt_list, 0) / (g(total_equity_list, 0) + g(total_debt_list, 0))
        e_ratio = 1 - d_ratio
        cost_of_equity = 8.0
        cost_of_debt = 3.0
        tax_rate = 0.25
        wacc_val = e_ratio * cost_of_equity + d_ratio * cost_of_debt * (1 - tax_rate)

    # SGA比率
    sga_ratio = None
    sga_ratio_5y = None
    if g(sga_list, 0) and g(revenue, 0):
        sga_ratio = (g(sga_list, 0) / g(revenue, 0)) * 100
    if g(sga_list, 4) and g(revenue, 4):
        sga_ratio_5y = (g(sga_list, 4) / g(revenue, 4)) * 100

    data = {
        "company": "",
        "ticker": "",
        "industry": "製造・サービス",

        "revenue": [g(revenue, i) for i in range(min(5, len(revenue)))],
        "fcf": [g(fcf_list, i) for i in range(min(5, len(fcf_list)))],
        "eps": [g(eps_list, i) for i in range(min(5, len(eps_list)))],

        "roe": [roe_now, roe_3y, roe_5y],
        "roe_growth_rate": roe_growth,
        "roa": [roa_now, roa_3y, roa_5y],

        "equity_ratio": equity_ratio,
        "equity_ratio_5y": equity_ratio_5y,
        "quick_ratio": quick_r,
        "quick_ratio_5y": quick_r_5y,
        "current_ratio": current_r,
        "current_ratio_5y": current_r_5y,

        "operating_cf": [g(ocf_list, i) for i in range(min(5, len(ocf_list)))],
        "investing_cf": [g(investing_cf_list, i) for i in range(min(5, len(investing_cf_list)))],
        "financing_cf": [g(financing_cf_list, i) for i in range(min(5, len(financing_cf_list)))],
        "op_margin": op_margin_vals,
        "ebitda_margin": ebitda_margin_val,
        "ebitda_margin_5y": ebitda_margin_5y,

        "debt_fcf": g(debt_fcf_list, 0),
        "debt_fcf_5y": g(debt_fcf_list, 4),
        "nd_ebitda": g(nd_ebitda_list, 0),
        "ev": g(ev_list, 0),
        "per": g(pe_list, 0),
        "per_5y": g(pe_list, 4),
        "pbr": g(pb_list, 0),
        "pbr_5y": g(pb_list, 4),

        "nopat": nopat,
        "nopat_5y": nopat_5y,
        "invested_capital": ic,
        "invested_capital_5y": ic_5y,
        "wacc": wacc_val,

        "accounts_receivable": g(receivables_list, 0),
        "accounts_receivable_5y": g(receivables_list, 4),
        "inventory": g(inventory_list, 0),
        "inventory_5y": g(inventory_list, 4),
        "accounts_payable": g(payables_list, 0),
        "accounts_payable_5y": g(payables_list, 4),
        "cogs": g(cogs_list, 0),
        "cogs_5y": g(cogs_list, 4),
        "sga_ratio": sga_ratio,
        "sga_ratio_5y": sga_ratio_5y,

        "total_assets": g(total_assets_list, 0),
        "total_assets_5y": g(total_assets_list, 4),
        "fixed_assets": g(fixed_assets_list, 0),
        "fixed_assets_5y": g(fixed_assets_list, 4),
        "tangible_fixed_assets": g(fixed_assets_list, 0),
        "tangible_fixed_assets_5y": g(fixed_assets_list, 4),
        "intangible_fixed_assets": g(intangibles_list, 0),
        "intangible_fixed_assets_5y": g(intangibles_list, 4),

        "net_income_val": g(net_income, 0),
        "net_income_val_5y": g(net_income, 4),
        "op_income_val": g(op_income, 0),
        "op_income_val_5y": g(op_income, 4),
        "interest_exp": g(interest_exp_list, 0),
        "interest_exp_5y": g(interest_exp_list, 4),
        "other_exp": g(other_exp_list, 0),
        "other_exp_5y": g(other_exp_list, 4),
        "pretax_income": g(pretax_income_list, 0),
        "pretax_income_5y": g(pretax_income_list, 4),
        "income_tax": g(income_tax_list, 0),
        "income_tax_5y": g(income_tax_list, 4),
        "total_equity": g(total_equity_list, 0),
        "total_equity_5y": g(total_equity_list, 4),

        "dividend_yield": to_pct(g(dividend_yield_list, 0)),
        "dividend_yield_5y": to_pct(g(dividend_yield_list, 4)),
        "payout_ratio": to_pct(g(payout_ratio_list, 0)),
        "payout_ratio_5y": to_pct(g(payout_ratio_list, 4)),

        "d1_mgmt_change": "○",
        "d2_ownership": "○",
        "d3_esg": "○",
    }

    # 時系列データ（チャート用）
    ts_data = {
        "dates": [str(d)[:4] if d else "" for d in dates_raw],
        "revenue": [g(revenue, i) for i in range(len(revenue))],
        "net_income": [g(net_income, i) for i in range(len(net_income))],
        "fcf": [g(fcf_list, i) for i in range(len(fcf_list))],
        "eps": [g(eps_list, i) for i in range(len(eps_list))],
        "ocf": [g(ocf_list, i) for i in range(len(ocf_list))],
        "investing_cf": [g(investing_cf_list, i) for i in range(len(investing_cf_list))],
        "financing_cf": [g(financing_cf_list, i) for i in range(len(financing_cf_list))],
        "ebitda": [g(ebitda_list, i) for i in range(len(ebitda_list))],
        "total_assets": [g(total_assets_list, i) for i in range(len(total_assets_list))],
        "total_equity": [g(total_equity_list, i) for i in range(len(total_equity_list))],
        "total_debt": [g(total_debt_list, i) for i in range(len(total_debt_list))],
        "roe": [to_pct(g(roe_list, i)) for i in range(len(roe_list))],
        "roa": [to_pct(g(roa_list, i)) for i in range(len(roa_list))],
        "op_margin": [to_pct(g(op_margin_list, i)) for i in range(len(op_margin_list))],
        "quick_ratio": [to_pct(g(quick_ratio_list, i)) for i in range(len(quick_ratio_list))],
        "current_ratio": [to_pct(g(current_ratio_list, i)) for i in range(len(current_ratio_list))],
        "equity_ratio": [((1 / (1 + g(debt_equity_list, i))) * 100) if (g(debt_equity_list, i) is not None) else ((g(total_equity_list, i) / g(total_assets_list, i) * 100) if (g(total_equity_list, i) and g(total_assets_list, i)) else None) for i in range(max(len(debt_equity_list), len(total_equity_list)))],
        "ebitda_margin": [to_pct(g(ebitda_margin_list, i)) for i in range(len(ebitda_margin_list))],
        "debt_fcf": [g(debt_fcf_list, i) for i in range(len(debt_fcf_list))],
        "roic": [to_pct(g(roic_list, i)) for i in range(len(roic_list))],
        "capex": [g(capex_list, i) for i in range(len(capex_list))],
        "sga": [g(sga_list, i) for i in range(len(sga_list))],
        "da": [g(da_list, i) for i in range(len(da_list))],
        "pe_ratio": [g(pe_list, i) for i in range(len(pe_list))],
        "pb_ratio": [g(pb_list, i) for i in range(len(pb_list))],
        "debt_ebitda": [g(debt_ebitda_list, i) for i in range(len(debt_ebitda_list))],
        "nd_ebitda": [g(nd_ebitda_list, i) for i in range(len(nd_ebitda_list))],
        "dividend_yield": [to_pct(g(dividend_yield_list, i)) for i in range(len(dividend_yield_list))],
        "payout_ratio": [to_pct(g(payout_ratio_list, i)) for i in range(len(payout_ratio_list))],
    }

    # 営業利益率の時系列が空で、計算可能な場合
    if not ts_data["op_margin"] and op_income and revenue:
        ts_data["op_margin"] = []
        for i in range(len(op_income)):
            oi = g(op_income, i)
            rev = g(revenue, i)
            if oi is not None and rev and rev != 0:
                ts_data["op_margin"].append(round(oi / rev * 100, 2))
            else:
                ts_data["op_margin"].append(None)

    # DuPont分解の時系列を計算: ROE = 純利益率 × 総資産回転率 × 財務レバレッジ
    max_len = len(revenue) if revenue else 0
    net_margin_ts = []
    asset_turnover_ts = []
    fin_leverage_ts = []
    for i in range(max_len):
        ni = g(net_income, i)
        rev = g(revenue, i)
        ta = g(total_assets_list, i)
        eq = g(total_equity_list, i)
        # 純利益率 (%) = Net Income / Revenue * 100
        if ni is not None and rev and rev != 0:
            net_margin_ts.append(round(ni / rev * 100, 2))
        else:
            net_margin_ts.append(None)
        # 総資産回転率 (x) = Revenue / Total Assets
        if rev is not None and ta and ta != 0:
            asset_turnover_ts.append(round(rev / ta, 3))
        else:
            asset_turnover_ts.append(None)
        # 財務レバレッジ (x) = Total Assets / Equity
        if ta is not None and eq and eq != 0:
            fin_leverage_ts.append(round(ta / eq, 3))
        else:
            fin_leverage_ts.append(None)

    ts_data["net_margin"] = net_margin_ts
    ts_data["asset_turnover"] = asset_turnover_ts
    ts_data["financial_leverage"] = fin_leverage_ts

    # 純利益率の分解時系列: 金利負担率・営業外損益率・税引後利益率
    interest_burden_ts = []  # (営業利益+金利)/ 営業利益 → Pretax前の金利影響
    tax_burden_ts = []       # 純利益 / 税引前利益
    nonop_burden_ts = []     # 税引前利益 / (営業利益+金利)
    for i in range(max_len):
        oi = g(op_income, i)
        pt = g(pretax_income_list, i)
        ni = g(net_income, i)
        ie = g(interest_exp_list, i)
        oe = g(other_exp_list, i)

        # 金利負担率: 金利控除後 / 営業利益
        # 営業利益 + 金利(通常マイナスが費用) = 金利控除後利益
        if oi is not None and oi != 0 and ie is not None:
            interest_burden_ts.append(round((oi + ie) / oi * 100, 2))
        elif oi is not None and oi != 0 and pt is not None:
            # fallback: pretax/op_income includes both interest and other
            interest_burden_ts.append(None)
        else:
            interest_burden_ts.append(None)

        # 営業外損益率: 税引前利益 / (営業利益 + 金利)
        if oi is not None and ie is not None and (oi + ie) != 0 and pt is not None:
            nonop_burden_ts.append(round(pt / (oi + ie) * 100, 2))
        else:
            nonop_burden_ts.append(None)

        # 税引後利益率: 純利益 / 税引前利益
        if pt is not None and pt != 0 and ni is not None:
            tax_burden_ts.append(round(ni / pt * 100, 2))
        else:
            tax_burden_ts.append(None)

    ts_data["interest_burden"] = interest_burden_ts
    ts_data["nonop_burden"] = nonop_burden_ts
    ts_data["tax_burden"] = tax_burden_ts

    return data, ts_data


# ---------- カスタム分析: 可視化可能データのスキャン ----------

# メトリクス定義: key, Excelラベル, シート名, カテゴリ, 単位, 表示名(ja), 表示名(en)
METRIC_CATALOG = [
    {"key": "revenue",          "sheet": "inc", "label": "Revenue",                    "cat": "growth",      "unit": "百万", "ja": "売上高",              "en": "Revenue"},
    {"key": "op_income",        "sheet": "inc", "label": "Operating Income",            "cat": "profitability","unit": "百万", "ja": "営業利益",            "en": "Operating Income"},
    {"key": "ebitda",           "sheet": "inc", "label": "EBITDA",                      "cat": "profitability","unit": "百万", "ja": "EBITDA",              "en": "EBITDA"},
    {"key": "net_income",       "sheet": "inc", "label": "Net Income",                 "cat": "profitability","unit": "百万", "ja": "純利益",              "en": "Net Income"},
    {"key": "eps",              "sheet": "inc", "label": "EPS (Basic)",                "cat": "growth",      "unit": "円",   "ja": "EPS",                 "en": "EPS"},
    {"key": "op_margin",        "sheet": "inc", "label": "Operating Margin",           "cat": "profitability","unit": "%",    "ja": "営業利益率",          "en": "Operating Margin"},
    {"key": "ebitda_margin",    "sheet": "inc", "label": "EBITDA Margin",              "cat": "profitability","unit": "%",    "ja": "EBITDAマージン",      "en": "EBITDA Margin"},
    {"key": "sga",              "sheet": "inc", "label": "Selling, General & Admin",   "cat": "other",       "unit": "百万", "ja": "販管費",              "en": "SG&A"},
    {"key": "da",               "sheet": "inc", "label": "Depreciation & Amortization","cat": "other",       "unit": "百万", "ja": "減価償却費",          "en": "D&A"},
    {"key": "cogs",             "sheet": "inc", "label": "Cost of Revenue",            "cat": "other",       "unit": "百万", "ja": "売上原価",            "en": "COGS"},
    {"key": "fcf",              "sheet": "cf",  "label": "Free Cash Flow",             "cat": "health",      "unit": "百万", "ja": "FCF",                 "en": "FCF"},
    {"key": "ocf",              "sheet": "cf",  "label": "Operating Cash Flow",        "cat": "health",      "unit": "百万", "ja": "営業CF",              "en": "Operating CF"},
    {"key": "capex",            "sheet": "cf",  "label": "Capital Expenditures",       "cat": "other",       "unit": "百万", "ja": "設備投資",            "en": "CapEx"},
    {"key": "investing_cf",     "sheet": "cf",  "label": "Investing Cash Flow",        "cat": "health",      "unit": "百万", "ja": "投資CF",              "en": "Investing CF"},
    {"key": "financing_cf",     "sheet": "cf",  "label": "Financing Cash Flow",        "cat": "health",      "unit": "百万", "ja": "財務CF",              "en": "Financing CF"},
    {"key": "total_assets",     "sheet": "bs",  "label": "Total Assets",               "cat": "health",      "unit": "百万", "ja": "総資産",              "en": "Total Assets"},
    {"key": "total_equity",     "sheet": "bs",  "label": "Shareholders Equity",        "cat": "health",      "unit": "百万", "ja": "自己資本",            "en": "Equity"},
    {"key": "total_debt",       "sheet": "bs",  "label": "Total Debt",                 "cat": "health",      "unit": "百万", "ja": "有利子負債",          "en": "Total Debt"},
    {"key": "cash",             "sheet": "bs",  "label": "Cash & Cash Equivalents",    "cat": "health",      "unit": "百万", "ja": "現金",                "en": "Cash"},
    {"key": "receivables",      "sheet": "bs",  "label": "Receivables",                "cat": "other",       "unit": "百万", "ja": "売上債権",            "en": "Receivables"},
    {"key": "inventory",        "sheet": "bs",  "label": "Inventory",                  "cat": "other",       "unit": "百万", "ja": "棚卸資産",            "en": "Inventory"},
    {"key": "current_assets",   "sheet": "bs",  "label": "Total Current Assets",       "cat": "health",      "unit": "百万", "ja": "流動資産",            "en": "Current Assets"},
    {"key": "current_liab",     "sheet": "bs",  "label": "Total Current Liabilities",  "cat": "health",      "unit": "百万", "ja": "流動負債",            "en": "Current Liabilities"},
    {"key": "pe_ratio",         "sheet": "rat", "label": "PE Ratio",                   "cat": "valuation",   "unit": "x",    "ja": "PER",                 "en": "P/E Ratio"},
    {"key": "pb_ratio",         "sheet": "rat", "label": "PB Ratio",                   "cat": "valuation",   "unit": "x",    "ja": "PBR",                 "en": "P/B Ratio"},
    {"key": "roe",              "sheet": "rat", "label": "Return on Equity (ROE)",     "cat": "performance", "unit": "%",     "ja": "ROE",                 "en": "ROE"},
    {"key": "roa",              "sheet": "rat", "label": "Return on Assets (ROA)",     "cat": "performance", "unit": "%",     "ja": "ROA",                 "en": "ROA"},
    {"key": "roic",             "sheet": "rat", "label": "Return on Invested Capital (ROIC)","cat": "performance","unit": "%","ja": "ROIC",                "en": "ROIC"},
    {"key": "dividend_yield",   "sheet": "rat", "label": "Dividend Yield",                "cat": "performance", "unit": "%",     "ja": "配当利回り",          "en": "Dividend Yield"},
    {"key": "payout_ratio",     "sheet": "rat", "label": "Payout Ratio",                  "cat": "performance", "unit": "%",     "ja": "配当性向",            "en": "Payout Ratio"},
    {"key": "current_ratio",    "sheet": "rat", "label": "Current Ratio",              "cat": "health",      "unit": "%",     "ja": "流動比率",            "en": "Current Ratio"},
    {"key": "quick_ratio",      "sheet": "rat", "label": "Quick Ratio",                "cat": "health",      "unit": "%",     "ja": "当座比率",            "en": "Quick Ratio"},
    {"key": "debt_ebitda",      "sheet": "rat", "label": "Debt/EBITDA",                "cat": "health",      "unit": "x",     "ja": "Debt/EBITDA",         "en": "Debt/EBITDA"},
    {"key": "nd_ebitda",        "sheet": "rat", "label": "Net Debt/EBITDA",            "cat": "valuation",   "unit": "x",     "ja": "Net Debt/EBITDA",     "en": "Net Debt/EBITDA"},
    {"key": "debt_fcf",         "sheet": "rat", "label": "Debt/FCF",                   "cat": "health",      "unit": "x",     "ja": "Debt/FCF",            "en": "Debt/FCF"},
    {"key": "ev",               "sheet": "rat", "label": "Enterprise Value",           "cat": "valuation",   "unit": "百万",  "ja": "EV",                  "en": "Enterprise Value"},
    {"key": "fixed_assets",     "sheet": "bs",  "label": "Property, Plant & Equipment","cat": "other",       "unit": "百万",  "ja": "有形固定資産",        "en": "PP&E"},
    {"key": "intangibles",      "sheet": "bs",  "label": "Goodwill and Intangibles",   "cat": "other",       "unit": "百万",  "ja": "のれん・無形資産",    "en": "Goodwill & Intangibles"},
    {"key": "net_debt",         "sheet": "bs",  "label": "Net Cash (Debt)",            "cat": "health",      "unit": "百万",  "ja": "ネットキャッシュ",    "en": "Net Cash (Debt)"},
    {"key": "long_term_assets", "sheet": "bs",  "label": "Total Long-Term Assets",     "cat": "other",       "unit": "百万",  "ja": "固定資産",            "en": "Long-Term Assets"},
]


def scan_available_metrics(filepath):
    """Excelファイルをスキャンし、可視化可能なメトリクス一覧を返す"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet_map = {
        'inc': _find_sheet(wb, ['Income-Annual', 'Income Statement', 'Income', 'Export', 'income']),
        'bs':  _find_sheet(wb, ['Balance-Sheet-Annual', 'Balance Sheet', 'Balance', 'balance']),
        'cf':  _find_sheet(wb, ['Cash-Flow-Annual', 'Cash Flow', 'Cash Flow Statement', 'cashflow']),
        'rat': _find_sheet(wb, ['Ratios-Annual', 'Ratios', 'Financial Ratios', 'ratios']),
    }

    available = []
    for m in METRIC_CATALOG:
        ws = sheet_map.get(m['sheet'])
        if ws is None:
            continue
        data = _get_row_data(ws, m['label'])
        if data:
            numeric = [v for v in data if isinstance(v, (int, float))]
            if numeric:
                available.append({
                    "key": m['key'],
                    "ja": m['ja'],
                    "en": m['en'],
                    "cat": m['cat'],
                    "unit": m['unit'],
                    "data_points": len(numeric),
                    "latest_value": numeric[0],
                })
    return available


def extract_custom_timeseries(filepath, selected_keys):
    """選択されたメトリクスの時系列データを返す"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet_map = {
        'inc': _find_sheet(wb, ['Income-Annual', 'Income Statement', 'Income', 'Export', 'income']),
        'bs':  _find_sheet(wb, ['Balance-Sheet-Annual', 'Balance Sheet', 'Balance', 'balance']),
        'cf':  _find_sheet(wb, ['Cash-Flow-Annual', 'Cash Flow', 'Cash Flow Statement', 'cashflow']),
        'rat': _find_sheet(wb, ['Ratios-Annual', 'Ratios', 'Financial Ratios', 'ratios']),
    }

    # 日付はincシートから取得
    inc_ws = sheet_map.get('inc')
    dates = _try_labels(inc_ws, ['Date', 'Year Ending', 'Fiscal Year', 'Period']) if inc_ws else []
    date_strs = [str(d)[:4] if d else "" for d in dates]

    result = {"dates": date_strs}
    catalog_map = {m['key']: m for m in METRIC_CATALOG}

    for key in selected_keys:
        m = catalog_map.get(key)
        if not m:
            continue
        ws = sheet_map.get(m['sheet'])
        if ws is None:
            continue
        data = _get_row_data(ws, m['label'])
        is_pct = m['unit'] == '%'
        if is_pct:
            data = [v * 100 if isinstance(v, (int, float)) else None for v in data]
        result[key] = data

    return result
