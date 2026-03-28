"""
個別株式分析エンジン
stock_analyzer.py のロジックをWebアプリ用にモジュール化
"""

DEFAULT_THRESHOLDS = {
    "equity_ratio_x": 20, "equity_ratio_tri": 35,
    "ebitda_margin_x": 5, "ebitda_margin_tri": 10,
    "op_margin_x": 3, "op_margin_tri": 5,
    "nd_ebitda_x": 5, "nd_ebitda_tri": 3,
    "per_hi": 40, "per_lo": 5,
    "pbr_hi": 5, "pbr_lo": 0.5,
    "current_ratio_x": 100, "current_ratio_tri": 120,
    "revenue_cagr_x": 0, "revenue_cagr_tri": 5,
    "eps_growth_x": 0, "eps_growth_tri": 10,
}

INDUSTRY_LIST = ["製造・サービス"]


def generate_dynamic_thresholds(benchmark):
    """Damodaranの業界平均データから動的に閾値を生成する。

    考え方:
    - 「業界平均の30%未満 → ×」「業界平均の60%未満 → ▲」（高い方が良い指標）
    - 「業界平均の2.5倍超 → ×」「業界平均の1.5倍超 → ▲」（低い方が良い指標）
    - PER/PBR はレンジ判定: 業界平均を中心に上下に幅を持たせる
    """
    if not benchmark:
        return DEFAULT_THRESHOLDS.copy()

    th = DEFAULT_THRESHOLDS.copy()

    # 自己資本比率: Damodaranの debt_to_capital_book から算出
    dtc = benchmark.get('debt_to_capital_book')
    if dtc is not None and dtc < 1:
        avg_eq = (1 - dtc) * 100  # 業界平均の自己資本比率(%)
        th["equity_ratio_x"] = round(avg_eq * 0.3, 1)
        th["equity_ratio_tri"] = round(avg_eq * 0.6, 1)

    # EBITDAマージン（業界平均が1%未満の場合はデフォルトのまま）
    em = benchmark.get('ebitda_margin')
    if em is not None and em > 0.01:
        avg_em = em * 100
        th["ebitda_margin_x"] = round(avg_em * 0.3, 1)
        th["ebitda_margin_tri"] = round(avg_em * 0.6, 1)

    # Net Debt / EBITDA: 低い方が良い → 業界平均の1.5倍で▲、2.5倍で×
    # 銀行等EBITDAが極小の業種ではDebt/EBITDAが異常値(>100)になるためスキップ
    nd = benchmark.get('debt_to_ebitda')
    if nd is not None and 0 < nd < 100:
        th["nd_ebitda_tri"] = round(nd * 1.5, 1)
        th["nd_ebitda_x"] = round(nd * 2.5, 1)

    # PER: 業界平均(Aggregate)を中心にレンジ
    pe = benchmark.get('pe_aggregate_all')
    if pe is not None and pe > 0:
        th["per_lo"] = round(pe * 0.2, 1)
        th["per_hi"] = round(pe * 2.5, 1)

    # PBR: 業界平均を中心にレンジ
    pbr = benchmark.get('pbr')
    if pbr is not None and pbr > 0:
        th["pbr_lo"] = round(pbr * 0.15, 2)
        th["pbr_hi"] = round(pbr * 2.5, 2)

    # 営業利益率: 業界平均の30%未満→×、60%未満→▲（平均<1%時はデフォルト）
    op = benchmark.get('operating_margin')
    if op is not None and op > 0.01:
        avg_op = op * 100
        th["op_margin_x"] = round(avg_op * 0.3, 1)
        th["op_margin_tri"] = round(avg_op * 0.6, 1)

    # 売上高CAGR: 業界の期待成長率を基準に
    eg = benchmark.get('expected_growth_5y')
    if eg is not None:
        avg_growth = eg * 100  # %に変換
        # ×: 業界期待成長率の-50%未満（衰退）、▲: 業界平均の50%未満
        th["revenue_cagr_x"] = round(min(avg_growth * -0.5, 0), 1)
        th["revenue_cagr_tri"] = round(max(avg_growth * 0.5, 0), 1)

    # EPS成長率: 売上高CAGRと同じ基準を流用
    if eg is not None:
        avg_growth = eg * 100
        th["eps_growth_x"] = round(min(avg_growth * -0.5, 0), 1)
        th["eps_growth_tri"] = round(max(avg_growth * 0.5, 0), 1)

    # 流動比率: Damodaranにデータなし → デフォルトのまま
    # ただし金融業界は構造的に低いため、debt_to_capitalが高い業界は緩和
    if dtc is not None and dtc > 0.6:
        th["current_ratio_x"] = 60
        th["current_ratio_tri"] = 80

    return th


def safe_div(a, b):
    try:
        if b is None or b == 0:
            return None
        return a / b
    except (TypeError, ZeroDivisionError):
        return None


def rate_change(current, base):
    r = safe_div(current, base)
    if r is None:
        return None
    return (r - 1) * 100


def consecutive_increase(vals):
    clean = [v for v in vals if v is not None]
    if len(clean) < 2:
        return None
    return all(clean[i] > clean[i + 1] for i in range(len(clean) - 1))


def analyze_quantitative(d):
    results = {}
    rev = d.get("revenue")
    fcf = d.get("fcf")
    eps = d.get("eps")
    roe = d.get("roe")
    roa = d.get("roa")
    equity_ratio = d.get("equity_ratio")
    quick_ratio = d.get("quick_ratio")
    current_ratio = d.get("current_ratio")
    op_cf = d.get("operating_cf")
    op_margin = d.get("op_margin")
    ebitda_margin = d.get("ebitda_margin")
    debt_fcf = d.get("debt_fcf")
    ev = d.get("ev")
    nopat = d.get("nopat")
    nopat_5y = d.get("nopat_5y")
    invested_capital = d.get("invested_capital")
    invested_capital_5y = d.get("invested_capital_5y")
    wacc = d.get("wacc")
    ar = d.get("accounts_receivable")
    ar_5y = d.get("accounts_receivable_5y")
    inventory = d.get("inventory")
    inventory_5y = d.get("inventory_5y")
    ap = d.get("accounts_payable")
    ap_5y = d.get("accounts_payable_5y")
    rev_5y = rev[3] if rev and len(rev) > 3 else None
    cogs = d.get("cogs")
    cogs_5y = d.get("cogs_5y")

    # 売上高推移
    if rev and len(rev) >= 4:
        r5 = rate_change(rev[0], rev[3])
        r3 = rate_change(rev[0], rev[2])
        c3 = consecutive_increase([rev[0], rev[1], rev[2]])
        c5 = consecutive_increase([rev[0], rev[1], rev[2], rev[3]])
        if r5 is None:
            ev_rev = "×"
        elif r5 >= 10 and c5:
            ev_rev = "◎"
        elif r5 >= 5:
            ev_rev = "○"
        elif r5 >= 0:
            ev_rev = "▲"
        else:
            ev_rev = "×"
        results["売上高推移"] = {
            "最新値": rev[0], "5年変化率": r5, "3年変化率": r3,
            "3年連続増加": c3, "5年連続増加": c5, "評価": ev_rev,
        }
    else:
        results["売上高推移"] = {"評価": "未入力"}

    # FCF推移
    if fcf and len(fcf) >= 4:
        f5 = rate_change(fcf[0], fcf[3])
        f3 = rate_change(fcf[0], fcf[2])
        c3 = consecutive_increase([fcf[0], fcf[1], fcf[2]])
        c5 = consecutive_increase([fcf[0], fcf[1], fcf[2], fcf[3]])
        if f5 is None:
            ev_fcf = "×"
        elif f5 >= 10 and c5:
            ev_fcf = "◎"
        elif f5 >= 5:
            ev_fcf = "○"
        elif f5 >= 0:
            ev_fcf = "▲"
        else:
            ev_fcf = "×"
        results["FCF推移"] = {
            "最新値": fcf[0], "5年変化率": f5, "3年変化率": f3,
            "3年連続増加": c3, "5年連続増加": c5, "評価": ev_fcf,
        }
    else:
        results["FCF推移"] = {"評価": "未入力"}

    # EPS推移
    if eps and len(eps) >= 4:
        e5 = rate_change(eps[0], eps[3])
        e3 = rate_change(eps[0], eps[2])
        c3 = consecutive_increase([eps[0], eps[1], eps[2]])
        c5 = consecutive_increase([eps[0], eps[1], eps[2], eps[3]])
        if e5 is None:
            ev_eps = "×"
        elif e5 >= 10 and c5:
            ev_eps = "◎"
        elif e5 >= 5:
            ev_eps = "○"
        elif e5 >= 0:
            ev_eps = "▲"
        else:
            ev_eps = "×"
        results["EPS推移"] = {
            "最新値": eps[0], "5年変化率": e5, "3年変化率": e3,
            "3年連続増加": c3, "5年連続増加": c5, "評価": ev_eps,
        }
    else:
        results["EPS推移"] = {"評価": "未入力"}

    # Debt/FCF
    if debt_fcf is not None:
        ev_dfc = "○" if debt_fcf < 3 else ("▲" if debt_fcf < 5 else "×")
        results["Debt/FCF"] = {"現在値": debt_fcf, "評価": ev_dfc}
    else:
        results["Debt/FCF"] = {"評価": "未入力"}

    # EV/FCF
    fcf_latest = fcf[0] if fcf else None
    ev_fcf_ratio = safe_div(ev, fcf_latest)
    if ev_fcf_ratio is not None:
        ev_efcf = "○" if ev_fcf_ratio < 15 else ("▲" if ev_fcf_ratio < 25 else "×")
        results["EV/FCF"] = {"現在値": ev_fcf_ratio, "評価": ev_efcf}
    else:
        results["EV/FCF"] = {"評価": "未入力"}

    # ROE
    if roe and len(roe) >= 3:
        roe_now, roe_3y, roe_5y_val = roe[0], roe[1], roe[2]
        roe_growth = d.get("roe_growth_rate")
        c3 = roe_now > roe_3y if (roe_now is not None and roe_3y is not None) else None
        c5 = roe_now > roe_5y_val if (roe_now is not None and roe_5y_val is not None) else None
        rg = roe_growth if roe_growth is not None else (roe_now - roe_5y_val if roe_now and roe_5y_val else None)
        if roe_now is None:
            ev_roe = "未入力"
        elif roe_now >= 15 and rg is not None and rg >= 5 and c3:
            ev_roe = "◎"
        elif roe_now >= 10 and rg is not None and rg >= 3:
            ev_roe = "○"
        elif roe_now >= 0 and rg is not None and rg >= 0:
            ev_roe = "▲"
        else:
            ev_roe = "×"
        results["ROE"] = {
            "現在値": roe_now,
            "3年変化pt": roe_now - roe_3y if roe_3y is not None else None,
            "5年変化pt": roe_now - roe_5y_val if roe_5y_val is not None else None,
            "評価": ev_roe,
        }
    else:
        results["ROE"] = {"評価": "未入力"}

    # ROA
    if roa and len(roa) >= 3:
        roa_now, roa_3y, roa_5y_val = roa[0], roa[1], roa[2]
        results["ROA"] = {
            "現在値": roa_now,
            "3年変化pt": roa_now - roa_3y if roa_3y is not None else None,
            "5年変化pt": roa_now - roa_5y_val if roa_5y_val is not None else None,
        }
    else:
        results["ROA"] = {"評価": "未入力"}

    # 配当利回り
    div_yield = d.get("dividend_yield")
    div_yield_5y = d.get("dividend_yield_5y")
    if div_yield is not None:
        ev_dy = "◎" if div_yield >= 4 else ("○" if div_yield >= 2 else ("▲" if div_yield >= 1 else "×"))
        dy_chg = div_yield - div_yield_5y if div_yield_5y is not None else None
        results["配当利回り"] = {"現在値": div_yield, "5年変化pt": dy_chg, "評価": ev_dy}
    else:
        results["配当利回り"] = {"評価": "未入力"}

    # 配当性向
    payout = d.get("payout_ratio")
    payout_5y = d.get("payout_ratio_5y")
    if payout is not None:
        ev_po = "◎" if 20 <= payout <= 50 else ("○" if 10 <= payout <= 70 else ("▲" if 0 <= payout <= 100 else "×"))
        po_chg = payout - payout_5y if payout_5y is not None else None
        results["配当性向"] = {"現在値": payout, "5年変化pt": po_chg, "評価": ev_po}
    else:
        results["配当性向"] = {"評価": "未入力"}

    # 自己資本比率
    equity_ratio_5y = d.get("equity_ratio_5y")
    if equity_ratio is not None:
        ev_eq = "◎" if equity_ratio >= 50 else ("○" if equity_ratio >= 40 else ("▲" if equity_ratio >= 20 else "×"))
        eq_chg = equity_ratio - equity_ratio_5y if equity_ratio_5y is not None else None
        results["自己資本比率"] = {"現在値": equity_ratio, "5年変化pt": eq_chg, "評価": ev_eq}
    else:
        results["自己資本比率"] = {"評価": "未入力"}

    # 当座比率
    quick_ratio_5y = d.get("quick_ratio_5y")
    if quick_ratio is not None:
        ev_qr = "◎" if quick_ratio >= 150 else ("○" if quick_ratio >= 100 else ("▲" if quick_ratio >= 80 else "×"))
        qr_chg = quick_ratio - quick_ratio_5y if quick_ratio_5y is not None else None
        results["当座比率"] = {"現在値": quick_ratio, "5年変化pt": qr_chg, "評価": ev_qr}
    else:
        results["当座比率"] = {"評価": "未入力"}

    # 流動比率
    current_ratio_5y = d.get("current_ratio_5y")
    if current_ratio is not None:
        ev_cr = "◎" if current_ratio >= 200 else ("○" if current_ratio >= 100 else ("▲" if current_ratio >= 80 else "×"))
        cr_chg = current_ratio - current_ratio_5y if current_ratio_5y is not None else None
        results["流動比率"] = {"現在値": current_ratio, "5年変化pt": cr_chg, "評価": ev_cr}
    else:
        results["流動比率"] = {"評価": "未入力"}

    # 営業CF
    if op_cf and len(op_cf) >= 1:
        latest_cf = op_cf[0]
        ev_cf = "○" if latest_cf and latest_cf > 0 else "×"
        cf_result = {"最新値": latest_cf, "評価": ev_cf}
        if len(op_cf) >= 4:
            cf_5y = rate_change(op_cf[0], op_cf[3])
            cf_result["5年変化率"] = cf_5y
        results["営業CF"] = cf_result
    else:
        results["営業CF"] = {"評価": "未入力"}

    # 投資CF
    inv_cf = d.get("investing_cf", [])
    if inv_cf and len(inv_cf) >= 1:
        latest_inv = inv_cf[0]
        ev_inv = "○" if latest_inv is not None and latest_inv < 0 else "▲"
        inv_result = {"最新値": latest_inv, "評価": ev_inv}
        if len(inv_cf) >= 4:
            inv_result["5年変化率"] = rate_change(abs(inv_cf[0]) if inv_cf[0] else None, abs(inv_cf[3]) if inv_cf[3] else None)
        results["投資CF"] = inv_result
    else:
        results["投資CF"] = {"評価": "未入力"}

    # 財務CF
    fin_cf = d.get("financing_cf", [])
    if fin_cf and len(fin_cf) >= 1:
        latest_fin = fin_cf[0]
        ev_fin = "○" if latest_fin is not None and latest_fin < 0 else "▲"
        fin_result = {"最新値": latest_fin, "評価": ev_fin}
        if len(fin_cf) >= 4:
            fin_result["5年変化率"] = rate_change(abs(fin_cf[0]) if fin_cf[0] else None, abs(fin_cf[3]) if fin_cf[3] else None)
        results["財務CF"] = fin_result
    else:
        results["財務CF"] = {"評価": "未入力"}

    # CF構成分析 (営業+, 投資-, 財務- = 理想型)
    ocf_ok = op_cf and len(op_cf) >= 1 and op_cf[0] is not None and op_cf[0] > 0
    icf_ok = inv_cf and len(inv_cf) >= 1 and inv_cf[0] is not None and inv_cf[0] < 0
    fcf_ok = fin_cf and len(fin_cf) >= 1 and fin_cf[0] is not None and fin_cf[0] < 0
    cf_pattern_ideal = ocf_ok and icf_ok and fcf_ok
    results["CF構成分析"] = {
        "営業CF符号": "+" if ocf_ok else ("-" if op_cf and len(op_cf)>=1 and op_cf[0] is not None else "N/A"),
        "投資CF符号": "-" if icf_ok else ("+" if inv_cf and len(inv_cf)>=1 and inv_cf[0] is not None else "N/A"),
        "財務CF符号": "-" if fcf_ok else ("+" if fin_cf and len(fin_cf)>=1 and fin_cf[0] is not None else "N/A"),
        "理想型": cf_pattern_ideal,
        "評価": "◎" if cf_pattern_ideal else "⚠",
    }

    # ROIC
    roic = safe_div(nopat, invested_capital)
    roic_5y = safe_div(nopat_5y, invested_capital_5y)
    if roic is not None:
        roic_pct = roic * 100
        ev_roic = "◎" if roic_pct >= 15 else ("○" if roic_pct >= 10 else ("▲" if roic_pct >= 0 else "×"))
        rev_latest = rev[0] if rev else None
        ic_turnover = safe_div(rev_latest, invested_capital)
        roic_chg = (roic - roic_5y) * 100 if roic_5y is not None else None
        results["ROIC"] = {
            "ROIC": roic_pct,
            "投下資本回転率": ic_turnover,
            "ROIC_5年変化pt": roic_chg,
            "評価": ev_roic,
        }
    else:
        results["ROIC"] = {"評価": "未入力"}

    # EBITDAマージン推移
    ebitda_margin_5y = d.get("ebitda_margin_5y")
    if ebitda_margin is not None:
        em_chg = ebitda_margin - ebitda_margin_5y if ebitda_margin_5y is not None else None
        results["EBITDAマージン"] = {"現在値": ebitda_margin, "5年変化pt": em_chg}

    # 営業利益率推移
    if op_margin and len(op_margin) >= 1:
        opm_now = op_margin[0]
        opm_5y = op_margin[4] if len(op_margin) >= 5 else None
        opm_chg = opm_now - opm_5y if (opm_now is not None and opm_5y is not None) else None
        results["営業利益率Q"] = {"現在値": opm_now, "5年変化pt": opm_chg}

    # Debt/FCF推移
    debt_fcf_5y = d.get("debt_fcf_5y")
    if debt_fcf is not None:
        df_chg = debt_fcf - debt_fcf_5y if debt_fcf_5y is not None else None
        results["Debt/FCF_detail"] = {"現在値": debt_fcf, "5年変化": df_chg}

    # WACC
    if wacc is not None:
        ev_wacc = "◎" if wacc <= 6 else ("○" if wacc <= 9 else ("▲" if wacc <= 12 else "×"))
        results["WACC"] = {"現在値": wacc, "評価": ev_wacc}
    else:
        results["WACC"] = {"評価": "未入力"}

    # ROIC vs WACC
    if roic is not None and wacc is not None:
        spread = roic * 100 - wacc
        ev_spread = "○" if spread > 0 else ("▲" if spread > wacc * (-0.2) else "×")
        results["ROIC_vs_WACC"] = {"スプレッド": spread, "評価": ev_spread}

    # CCC分析
    rev_now = rev[0] if rev else None
    if all(v is not None for v in [ar, inventory, ap, rev_now, cogs]):
        dso = safe_div(ar, rev_now) * 365
        dio = safe_div(inventory, cogs) * 365
        dpo = safe_div(ap, rev_now) * 365
        ccc = dso + dio - dpo if (dso is not None and dio is not None and dpo is not None) else None
        dso_5y_v = safe_div(ar_5y, rev_5y) * 365 if (ar_5y and rev_5y) else None
        dio_5y_v = safe_div(inventory_5y, cogs_5y) * 365 if (inventory_5y and cogs_5y) else None
        dpo_5y_v = safe_div(ap_5y, rev_5y) * 365 if (ap_5y and rev_5y) else None
        ccc_5y = dso_5y_v + dio_5y_v - dpo_5y_v if (dso_5y_v is not None and dio_5y_v is not None and dpo_5y_v is not None) else None
        ev_ccc = "◎" if (ccc and ccc < 30) else ("○" if (ccc and ccc < 60) else ("▲" if (ccc and ccc < 90) else "×"))
        results["CCC分析"] = {
            "CCC": ccc, "CCC_5年変化": ccc - ccc_5y if ccc_5y else None,
            "DSO": dso, "DIO": dio, "DPO": dpo,
            "評価": ev_ccc,
        }
    else:
        results["CCC分析"] = {"評価": "未入力"}

    return results


def analyze_screening(d, q_results, benchmark=None):
    th = generate_dynamic_thresholds(benchmark)
    # 業界平均ROE（日本 or Global）
    avg_roe = None
    if benchmark:
        avg_roe = benchmark.get('roe_japan')
        if avg_roe is None:
            avg_roe = benchmark.get('roe_global')
        if avg_roe is not None:
            avg_roe = avg_roe * 100  # %に変換
    results = {}

    # Section A: 財務健全性
    eq_r = d.get("equity_ratio")
    a1 = "未入力" if eq_r is None else ("×" if eq_r < th["equity_ratio_x"] else ("▲" if eq_r < th["equity_ratio_tri"] else "○"))
    a1_basis = f"○≧{th['equity_ratio_tri']}% ▲≧{th['equity_ratio_x']}% ×<{th['equity_ratio_x']}%（業界平均の30%/60%基準）"
    results["A-1_自己資本比率"] = {"実績値": eq_r, "閾値_x": th["equity_ratio_x"], "閾値_tri": th["equity_ratio_tri"], "判定": a1, "基準": a1_basis}

    op_cf = d.get("operating_cf", [])
    if not op_cf:
        a2, a2_val = "未入力", None
    else:
        cf0, cf1, cf2 = (op_cf + [None, None, None])[:3]
        if cf0 is None:
            a2, a2_val = "未入力", None
        elif cf0 > 0:
            a2, a2_val = "○", "最新プラス"
        elif cf0 < 0 and cf1 is not None and cf1 < 0 and cf2 is not None and cf2 < 0:
            a2, a2_val = "×", "3期連続マイナス"
        elif cf0 < 0 and cf1 is not None and cf1 < 0:
            a2, a2_val = "▲", "前年2期マイナス"
        else:
            a2, a2_val = "▲", "約1期マイナス"
    a2_basis = "○=最新期プラス ▲=1-2期マイナス ×=3期連続マイナス"
    results["A-2_営業CF"] = {"状況": a2_val, "判定": a2, "基準": a2_basis}

    roe = d.get("roe", [])
    roa_d = d.get("roa", [])
    roe_now = roe[0] if roe else None
    roa_now = roa_d[0] if roa_d else None
    if roe_now is None:
        a3 = "未入力"
    elif avg_roe is not None:
        # Damodaran業界平均ROEとの比較
        if roe_now < 0:
            a3 = "×"
        elif roe_now < avg_roe * 0.5:
            a3 = "▲"
        else:
            a3 = "○"
    else:
        # フォールバック: 正負判定
        if roe_now is not None and roa_now is not None:
            if roe_now < 0 and roa_now < 0:
                a3 = "×"
            elif roe_now < 0 or roa_now < 0:
                a3 = "▲"
            else:
                a3 = "○"
        else:
            a3 = "未入力"
    if avg_roe is not None:
        a3_basis = f"○≧業界平均の50%({avg_roe*0.5:.1f}%) ▲<{avg_roe*0.5:.1f}% ×=マイナス（業界平均ROE: {avg_roe:.1f}%）"
    else:
        a3_basis = "○=ROEプラス ▲=一方マイナス ×=両方マイナス"
    results["A-3_ROE_ROA"] = {"ROE最新": roe_now, "ROA最新": roa_now, "判定": a3, "基準": a3_basis}

    a_keys = ["A-1_自己資本比率", "A-2_営業CF", "A-3_ROE_ROA"]
    a_xs = sum(1 for k in a_keys if results[k]["判定"] == "×")
    a_tris = sum(1 for k in a_keys if results[k]["判定"] == "▲")
    if a_xs >= 1:
        a_eval = "×NG（即時回避）"
    elif a_tris >= 2:
        a_eval = "▲要注意"
    elif a_tris >= 1:
        a_eval = "▲要確認"
    else:
        a_eval = "◎全項通過"
    a_score = (a_xs * 1 + a_tris * 0.5) * 5
    results["SectionA評価"] = a_eval
    results["SectionAスコア"] = a_score

    # Section B: 成長性・収益性
    rev = d.get("revenue", [])
    if len(rev) >= 3 and rev[0] and rev[2]:
        cagr3 = (rev[0] / rev[2]) ** (1 / 3) - 1
        cagr3_pct = cagr3 * 100
        b1 = "×" if cagr3_pct < th["revenue_cagr_x"] else ("▲" if cagr3_pct < th["revenue_cagr_tri"] else "○")
    else:
        cagr3_pct, b1 = None, "未入力"
    b1_basis = f"○≧{th['revenue_cagr_tri']}% ▲≧{th['revenue_cagr_x']}% ×<{th['revenue_cagr_x']}%（業界期待成長率の±50%基準）"
    results["B-1_売上高CAGR"] = {"3年CAGR": cagr3_pct, "判定": b1, "基準": b1_basis}

    em = d.get("ebitda_margin")
    b2 = "未入力" if em is None else ("×" if em < th["ebitda_margin_x"] else ("▲" if em < th["ebitda_margin_tri"] else "○"))
    b2_basis = f"○≧{th['ebitda_margin_tri']}% ▲≧{th['ebitda_margin_x']}% ×<{th['ebitda_margin_x']}%（業界平均の30%/60%基準）"
    results["B-2_EBITDAマージン"] = {"実績値": em, "閾値_x": th["ebitda_margin_x"], "閾値_tri": th["ebitda_margin_tri"], "判定": b2, "基準": b2_basis}

    op_m = d.get("op_margin", [])
    if len(op_m) >= 3:
        op_now, op_1y, op_2y = op_m[0], op_m[1], op_m[2]
        if op_now is not None and op_1y is not None and op_2y is not None:
            if op_now < op_1y and op_1y < op_2y:
                b3_val, b3 = "連続低下（3期）", "×"
            elif op_now < th["op_margin_x"]:
                b3_val, b3 = op_now, "×"
            elif op_now < th["op_margin_tri"]:
                b3_val, b3 = op_now, "▲"
            else:
                b3_val, b3 = op_now, "○"
        else:
            b3_val, b3 = None, "未入力"
    else:
        b3_val, b3 = None, "未入力"
    b3_basis = f"○≧{th['op_margin_tri']}% ▲≧{th['op_margin_x']}% ×<{th['op_margin_x']}%または3期連続低下（業界平均の30%/60%基準）"
    results["B-3_営業利益率"] = {"状況": b3_val, "判定": b3, "基準": b3_basis}

    eps = d.get("eps", [])
    if len(eps) >= 4 and eps[0] and eps[3]:
        eps_gr = (eps[0] / eps[3] - 1) * 100
        b4 = "×" if eps_gr < th["eps_growth_x"] else ("▲" if eps_gr < th["eps_growth_tri"] else "○")
    else:
        eps_gr, b4 = None, "未入力"
    b4_basis = f"○≧{th['eps_growth_tri']}% ▲≧{th['eps_growth_x']}% ×<{th['eps_growth_x']}%（業界期待成長率の±50%基準）"
    results["B-4_EPS成長率"] = {"5年成長率": eps_gr, "判定": b4, "基準": b4_basis}

    b_keys = ["B-1_売上高CAGR", "B-2_EBITDAマージン", "B-3_営業利益率", "B-4_EPS成長率"]
    b_xs = sum(1 for k in b_keys if results[k]["判定"] == "×")
    b_tris = sum(1 for k in b_keys if results[k]["判定"] == "▲")
    b_eval = "×投資魅力低下" if b_xs >= 2 else ("▲要追加調査" if b_xs >= 1 else ("▲要注意" if b_tris >= 2 else "◎全項通過"))
    b_score = (b_xs * 1 + b_tris * 0.5) * 4
    results["SectionB評価"] = b_eval
    results["SectionBスコア"] = b_score

    # Section C: バリュエーション・リスク
    per = d.get("per")
    pbr = d.get("pbr")
    nd_ebitda = d.get("nd_ebitda")
    cur_r = d.get("current_ratio")

    c1 = "未入力" if per is None else ("×" if (per > th["per_hi"] or per < th["per_lo"]) else ("▲" if (per > th["per_hi"] * 0.8 or per < th["per_lo"] * 1.5) else "○"))
    c1_basis = f"○={th['per_lo']}～{th['per_hi']}x ▲=やや範囲外 ×<{th['per_lo']}xまたは>{th['per_hi']}x（業界平均PERの0.2～2.5倍基準）"
    results["C-1_PER"] = {"実績値": per, "閾値_上限": th["per_hi"], "閾値_下限": th["per_lo"], "判定": c1, "基準": c1_basis}

    c2 = "未入力" if pbr is None else ("×" if (pbr > th["pbr_hi"] or pbr < th["pbr_lo"]) else ("▲" if (pbr > th["pbr_hi"] * 0.7 or pbr < th["pbr_lo"] * 1.5) else "○"))
    c2_basis = f"○={th['pbr_lo']}～{th['pbr_hi']}x ▲=やや範囲外 ×<{th['pbr_lo']}xまたは>{th['pbr_hi']}x（業界平均PBRの0.15～2.5倍基準）"
    results["C-2_PBR"] = {"実績値": pbr, "閾値_上限": th["pbr_hi"], "閾値_下限": th["pbr_lo"], "判定": c2, "基準": c2_basis}

    c3 = "未入力" if nd_ebitda is None else ("×" if nd_ebitda > th["nd_ebitda_x"] else ("▲" if nd_ebitda > th["nd_ebitda_tri"] else "○"))
    c3_basis = f"○≦{th['nd_ebitda_tri']}x ▲≦{th['nd_ebitda_x']}x ×>{th['nd_ebitda_x']}x（業界平均の1.5/2.5倍基準）"
    results["C-3_NetDebt_EBITDA"] = {"実績値": nd_ebitda, "閾値_x": th["nd_ebitda_x"], "閾値_tri": th["nd_ebitda_tri"], "判定": c3, "基準": c3_basis}

    c4 = "未入力" if cur_r is None else ("×" if cur_r < th["current_ratio_x"] else ("▲" if cur_r < th["current_ratio_tri"] else "○"))
    c4_basis = f"○≧{th['current_ratio_tri']}% ▲≧{th['current_ratio_x']}% ×<{th['current_ratio_x']}%"
    results["C-4_流動比率"] = {"実績値": cur_r, "閾値_x": th["current_ratio_x"], "閾値_tri": th["current_ratio_tri"], "判定": c4, "基準": c4_basis}

    c_keys = ["C-1_PER", "C-2_PBR", "C-3_NetDebt_EBITDA", "C-4_流動比率"]
    c_xs = sum(1 for k in c_keys if results[k]["判定"] == "×")
    c_tris = sum(1 for k in c_keys if results[k]["判定"] == "▲")
    c_eval = "×タイミング要見直し" if c_xs >= 3 else ("▲要コスト評価" if c_xs >= 1 else ("▲要注意" if c_tris >= 2 else "◎全項通過"))
    c_score = (c_xs * 1 + c_tris * 0.5) * 3
    results["SectionC評価"] = c_eval
    results["SectionCスコア"] = c_score

    # Section D: 定性・ESG
    d1 = d.get("d1_mgmt_change", "未入力")
    d2 = d.get("d2_ownership", "未入力")
    d3 = d.get("d3_esg", "未入力")
    d1_basis = "○=安定経営陣 ▲=一部変更 ×=大幅刷新・不祥事"
    results["D-1_経営陣変更"] = {"判定": d1, "基準": d1_basis}
    d2_basis = "○=安定株主構成 ▲=一部変動 ×=敵対的買収・大量売却"
    results["D-2_株主構造"] = {"判定": d2, "基準": d2_basis}
    d3_basis = "○=ESGリスクなし ▲=軽微なリスク ×=重大なESG・規制リスク"
    results["D-3_ESG"] = {"判定": d3, "基準": d3_basis}

    d_xs = sum(1 for v in [d1, d2, d3] if v == "×")
    d_tris = sum(1 for v in [d1, d2, d3] if v == "▲")
    d_eval = "×重要リスク確認" if d_xs >= 2 else ("▲リスク注意" if d_xs >= 1 else ("▲要確認" if d_tris >= 1 else "◎リスクなし"))
    d_score = (d_xs * 1 + d_tris * 0.5) * 2
    results["SectionD評価"] = d_eval
    results["SectionDスコア"] = d_score

    # 最終投資判定
    total_score = a_score + b_score + c_score + d_score
    if a_xs >= 1 or total_score >= 15:
        final = "SELL"
    elif total_score <= 5:
        final = "BUY"
    else:
        final = "HOLD"
    results["総合スコア"] = total_score
    results["最終投資判定"] = final

    return results


def analyze_roa_tree(d):
    rev = d.get("revenue", [])
    roa = d.get("roa", [])
    op_margin = d.get("op_margin", [])
    rev_now = rev[0] if rev else None
    rev_5y = rev[3] if len(rev) > 3 else None
    roa_now = roa[0] if roa else None
    roa_5y = roa[2] if len(roa) > 2 else None
    op_now = op_margin[0] if op_margin else None
    op_5y = op_margin[2] if len(op_margin) > 2 else None
    total_assets = d.get("total_assets")
    total_assets_5y = d.get("total_assets_5y")
    fixed_assets = d.get("fixed_assets")
    fixed_assets_5y = d.get("fixed_assets_5y")
    tangible_assets = d.get("tangible_fixed_assets")
    tangible_5y = d.get("tangible_fixed_assets_5y")
    intangible_assets = d.get("intangible_fixed_assets")
    intangible_5y = d.get("intangible_fixed_assets_5y")
    ar = d.get("accounts_receivable")
    ar_5y = d.get("accounts_receivable_5y")
    ap = d.get("accounts_payable")
    ap_5y = d.get("accounts_payable_5y")
    inventory = d.get("inventory")
    inventory_5y = d.get("inventory_5y")
    cogs = d.get("cogs")
    cogs_5y = d.get("cogs_5y")
    sga = d.get("sga_ratio")
    sga_5y = d.get("sga_ratio_5y")

    tree = {}

    tree["ROA"] = {
        "現在値": roa_now,
        "5年変化pt": roa_now - roa_5y if (roa_now and roa_5y) else None,
    }

    asset_turn = safe_div(rev_now, total_assets)
    asset_turn_5y = safe_div(rev_5y, total_assets_5y)
    tree["資産回転率"] = {
        "現在値": asset_turn,
        "5年変化": asset_turn - asset_turn_5y if (asset_turn and asset_turn_5y) else None,
    }

    fixed_turn = safe_div(rev_now, fixed_assets)
    fixed_turn_5y = safe_div(rev_5y, fixed_assets_5y)
    tree["固定資産回転率"] = {
        "現在値": fixed_turn,
        "5年変化": fixed_turn - fixed_turn_5y if (fixed_turn and fixed_turn_5y) else None,
    }

    tang_turn = safe_div(rev_now, tangible_assets)
    tang_turn_5y = safe_div(rev_5y, tangible_5y)
    tang_chg = tang_turn - tang_turn_5y if (tang_turn and tang_turn_5y) else None
    tree["有形固定資産回転率"] = {"現在値": tang_turn, "5年変化": tang_chg, "評価": "改善" if tang_chg and tang_chg > 0 else ("悪化" if tang_chg and tang_chg < 0 else "横ばい")}

    intan_turn = safe_div(rev_now, intangible_assets)
    intan_turn_5y = safe_div(rev_5y, intangible_5y)
    intan_chg = intan_turn - intan_turn_5y if (intan_turn and intan_turn_5y) else None
    tree["無形固定資産回転率"] = {"現在値": intan_turn, "5年変化": intan_chg, "評価": "改善" if intan_chg and intan_chg > 0 else ("悪化" if intan_chg and intan_chg < 0 else "横ばい")}

    dso = safe_div(ar, rev_now) * 365 if (ar and rev_now) else None
    dso_5y_v = safe_div(ar_5y, rev_5y) * 365 if (ar_5y and rev_5y) else None
    dso_chg = -(dso - dso_5y_v) if (dso and dso_5y_v) else None

    dpo = safe_div(ap, rev_now) * 365 if (ap and rev_now) else None
    dpo_5y_v = safe_div(ap_5y, rev_5y) * 365 if (ap_5y and rev_5y) else None
    dpo_chg = dpo - dpo_5y_v if (dpo and dpo_5y_v) else None

    dio = safe_div(inventory, cogs) * 365 if (inventory and cogs) else None
    dio_5y_v = safe_div(inventory_5y, cogs_5y) * 365 if (inventory_5y and cogs_5y) else None
    dio_chg = -(dio - dio_5y_v) if (dio and dio_5y_v) else None

    tree["DSO"] = {"現在値": dso, "5年変化": dso_chg, "評価": "改善" if dso_chg and dso_chg > 0 else ("悪化" if dso_chg and dso_chg < 0 else "横ばい")}
    tree["DPO"] = {"現在値": dpo, "5年変化": dpo_chg, "評価": "改善" if dpo_chg and dpo_chg > 0 else ("悪化" if dpo_chg and dpo_chg < 0 else "横ばい")}
    tree["DIO"] = {"現在値": dio, "5年変化": dio_chg, "評価": "改善" if dio_chg and dio_chg > 0 else ("悪化" if dio_chg and dio_chg < 0 else "横ばい")}

    tree["営業利益率"] = {
        "現在値": op_now,
        "5年変化pt": op_now - op_5y if (op_now and op_5y) else None,
    }

    cogs_rate = safe_div(cogs, rev_now) * 100 if (cogs and rev_now) else None
    cogs_rate_5y = safe_div(cogs_5y, rev_5y) * 100 if (cogs_5y and rev_5y) else None
    cogs_chg = -(cogs_rate - cogs_rate_5y) if (cogs_rate and cogs_rate_5y) else None
    tree["原価率"] = {"現在値": cogs_rate, "5年変化": cogs_chg, "評価": "改善" if cogs_chg and cogs_chg > 0 else ("悪化" if cogs_chg and cogs_chg < 0 else "横ばい")}

    sga_chg = -(sga - sga_5y) if (sga and sga_5y) else None
    tree["販管費率"] = {"現在値": sga, "5年変化": sga_chg, "評価": "改善" if sga_chg and sga_chg > 0 else ("悪化" if sga_chg and sga_chg < 0 else "横ばい")}

    contributions = {
        "有形固定資産回転率": tang_chg,
        "無形固定資産回転率": intan_chg,
        "DSO": dso_chg, "DPO": dpo_chg, "DIO": dio_chg,
        "原価率": cogs_chg, "販管費率": sga_chg,
    }
    ranked = sorted(
        [(k, v) for k, v in contributions.items() if v is not None],
        key=lambda x: abs(x[1]), reverse=True
    )
    tree["貢献度ランキング"] = [{"順位": i + 1, "指標": k, "改善寄与度": v, "評価": "改善" if v > 0 else "悪化"} for i, (k, v) in enumerate(ranked)]

    return tree


def analyze_roe_tree(d):
    """DuPont ROE decomposition (3-factor)"""
    rev = d.get("revenue", [])
    roe_vals = d.get("roe", [])
    rev_now = rev[0] if rev else None
    rev_5y = rev[3] if len(rev) > 3 else None
    net_income = d.get("net_income_val")
    net_income_5y = d.get("net_income_val_5y")
    total_assets = d.get("total_assets")
    total_assets_5y = d.get("total_assets_5y")
    total_equity = d.get("total_equity")
    total_equity_5y = d.get("total_equity_5y")
    roe_now = roe_vals[0] if roe_vals else None
    roe_5y = roe_vals[2] if len(roe_vals) > 2 else None

    tree = {}
    tree["ROE"] = {
        "現在値": roe_now,
        "5年変化pt": roe_now - roe_5y if (roe_now is not None and roe_5y is not None) else None,
    }

    # Factor 1: Net Profit Margin = Net Income / Revenue
    npm = safe_div(net_income, rev_now)
    npm_pct = npm * 100 if npm is not None else None
    npm_5y = safe_div(net_income_5y, rev_5y)
    npm_5y_pct = npm_5y * 100 if npm_5y is not None else None
    npm_chg = npm_pct - npm_5y_pct if (npm_pct is not None and npm_5y_pct is not None) else None
    tree["純利益率"] = {
        "現在値": npm_pct,
        "5年前": npm_5y_pct,
        "5年変化pt": npm_chg,
        "評価": "改善" if npm_chg and npm_chg > 0 else ("悪化" if npm_chg and npm_chg < 0 else "横ばい"),
    }

    # Factor 2: Asset Turnover = Revenue / Total Assets
    at = safe_div(rev_now, total_assets)
    at_5y = safe_div(rev_5y, total_assets_5y)
    at_chg = at - at_5y if (at is not None and at_5y is not None) else None
    tree["総資産回転率"] = {
        "現在値": at,
        "5年前": at_5y,
        "5年変化": at_chg,
        "評価": "改善" if at_chg and at_chg > 0 else ("悪化" if at_chg and at_chg < 0 else "横ばい"),
    }

    # Factor 3: Equity Multiplier = Total Assets / Equity
    em = safe_div(total_assets, total_equity)
    em_5y = safe_div(total_assets_5y, total_equity_5y)
    em_chg = em - em_5y if (em is not None and em_5y is not None) else None
    tree["財務レバレッジ"] = {
        "現在値": em,
        "5年前": em_5y,
        "5年変化": em_chg,
        "評価": "横ばい",
    }

    # Verify: ROE ≈ NPM × AT × EM
    if npm is not None and at is not None and em is not None:
        computed_roe = npm * at * em * 100
        tree["ROE検算"] = {"計算値": computed_roe}

    # Sub-decomposition of Net Profit Margin
    # 純利益率 = 営業利益率 × 金利負担率 × 営業外損益率 × 税引後利益率
    op_margin = d.get("op_margin", [])
    op_now = op_margin[0] if op_margin else None
    op_5y = op_margin[2] if len(op_margin) > 2 else None
    cogs = d.get("cogs")
    cogs_5y = d.get("cogs_5y")
    sga = d.get("sga_ratio")
    sga_5y = d.get("sga_ratio_5y")

    cogs_rate = safe_div(cogs, rev_now) * 100 if (cogs and rev_now) else None
    cogs_rate_5y = safe_div(cogs_5y, rev_5y) * 100 if (cogs_5y and rev_5y) else None
    cogs_chg = -(cogs_rate - cogs_rate_5y) if (cogs_rate is not None and cogs_rate_5y is not None) else None
    tree["原価率"] = {"現在値": cogs_rate, "5年変化": cogs_chg, "評価": "改善" if cogs_chg and cogs_chg > 0 else ("悪化" if cogs_chg and cogs_chg < 0 else "横ばい")}

    sga_chg = -(sga - sga_5y) if (sga is not None and sga_5y is not None) else None
    tree["販管費率"] = {"現在値": sga, "5年変化": sga_chg, "評価": "改善" if sga_chg and sga_chg > 0 else ("悪化" if sga_chg and sga_chg < 0 else "横ばい")}

    tree["営業利益率"] = {
        "現在値": op_now,
        "5年変化pt": op_now - op_5y if (op_now is not None and op_5y is not None) else None,
    }

    # --- 営業利益率→純利益率のギャップ分解 ---
    op_income_val = d.get("op_income_val")
    op_income_val_5y = d.get("op_income_val_5y")
    interest_exp = d.get("interest_exp")
    interest_exp_5y = d.get("interest_exp_5y")
    other_exp = d.get("other_exp")
    other_exp_5y = d.get("other_exp_5y")
    pretax_income = d.get("pretax_income")
    pretax_income_5y = d.get("pretax_income_5y")
    income_tax = d.get("income_tax")
    income_tax_5y = d.get("income_tax_5y")

    # 金利負担率 (%) = (営業利益 + 金利収支) / 営業利益 × 100
    # 100%なら金利負担ゼロ、低いほど金利負担が重い
    int_burden = None
    int_burden_5y = None
    if op_income_val and op_income_val != 0 and interest_exp is not None:
        int_burden = (op_income_val + interest_exp) / op_income_val * 100
    if op_income_val_5y and op_income_val_5y != 0 and interest_exp_5y is not None:
        int_burden_5y = (op_income_val_5y + interest_exp_5y) / op_income_val_5y * 100
    int_burden_chg = int_burden - int_burden_5y if (int_burden is not None and int_burden_5y is not None) else None
    tree["金利負担率"] = {
        "現在値": round(int_burden, 2) if int_burden is not None else None,
        "5年前": round(int_burden_5y, 2) if int_burden_5y is not None else None,
        "5年変化pt": round(int_burden_chg, 2) if int_burden_chg is not None else None,
        "評価": "改善" if int_burden_chg and int_burden_chg > 0 else ("悪化" if int_burden_chg and int_burden_chg < 0 else "横ばい"),
        "説明": "100%=金利負担なし。低いほど金利コストが重い",
    }

    # 営業外損益率 (%) = 税引前利益 / (営業利益 + 金利収支) × 100
    # 100%なら営業外損益ゼロ、乖離は為替・資産売却等の一時要因
    nonop_burden = None
    nonop_burden_5y = None
    oi_after_int = (op_income_val + interest_exp) if (op_income_val and interest_exp is not None) else None
    oi_after_int_5y = (op_income_val_5y + interest_exp_5y) if (op_income_val_5y and interest_exp_5y is not None) else None
    if oi_after_int and oi_after_int != 0 and pretax_income is not None:
        nonop_burden = pretax_income / oi_after_int * 100
    if oi_after_int_5y and oi_after_int_5y != 0 and pretax_income_5y is not None:
        nonop_burden_5y = pretax_income_5y / oi_after_int_5y * 100
    nonop_chg = nonop_burden - nonop_burden_5y if (nonop_burden is not None and nonop_burden_5y is not None) else None
    tree["営業外損益率"] = {
        "現在値": round(nonop_burden, 2) if nonop_burden is not None else None,
        "5年前": round(nonop_burden_5y, 2) if nonop_burden_5y is not None else None,
        "5年変化pt": round(nonop_chg, 2) if nonop_chg is not None else None,
        "評価": "改善" if nonop_chg and nonop_chg > 0 else ("悪化" if nonop_chg and nonop_chg < 0 else "横ばい"),
        "説明": "100%=営業外損益なし。為替差損益・資産売却等の影響",
    }

    # 税引後利益率 (%) = 純利益 / 税引前利益 × 100 = (1 - 実効税率)
    tax_burden = None
    tax_burden_5y = None
    if pretax_income and pretax_income != 0 and net_income is not None:
        tax_burden = net_income / pretax_income * 100
    if pretax_income_5y and pretax_income_5y != 0 and net_income_5y is not None:
        tax_burden_5y = net_income_5y / pretax_income_5y * 100
    tax_chg = tax_burden - tax_burden_5y if (tax_burden is not None and tax_burden_5y is not None) else None
    tree["税引後利益率"] = {
        "現在値": round(tax_burden, 2) if tax_burden is not None else None,
        "5年前": round(tax_burden_5y, 2) if tax_burden_5y is not None else None,
        "5年変化pt": round(tax_chg, 2) if tax_chg is not None else None,
        "評価": "改善" if tax_chg and tax_chg > 0 else ("悪化" if tax_chg and tax_chg < 0 else "横ばい"),
        "説明": "100%=税金ゼロ。低いほど税負担が重い（≒1-実効税率）",
    }

    # 営業利益→純利益ギャップの検算
    if op_now is not None and int_burden is not None and nonop_burden is not None and tax_burden is not None:
        computed_npm = op_now * (int_burden / 100) * (nonop_burden / 100) * (tax_burden / 100)
        tree["純利益率検算"] = {
            "計算値": round(computed_npm, 2),
            "実績値": round(npm_pct, 2) if npm_pct is not None else None,
        }

    # Sub-decomposition of Asset Turnover
    fixed_assets = d.get("fixed_assets")
    fixed_assets_5y = d.get("fixed_assets_5y")
    ar = d.get("accounts_receivable")
    ar_5y = d.get("accounts_receivable_5y")
    inventory = d.get("inventory")
    inventory_5y = d.get("inventory_5y")

    fixed_turn = safe_div(rev_now, fixed_assets)
    fixed_turn_5y = safe_div(rev_5y, fixed_assets_5y)
    fixed_chg = fixed_turn - fixed_turn_5y if (fixed_turn is not None and fixed_turn_5y is not None) else None
    tree["固定資産回転率"] = {"現在値": fixed_turn, "5年変化": fixed_chg, "評価": "改善" if fixed_chg and fixed_chg > 0 else ("悪化" if fixed_chg and fixed_chg < 0 else "横ばい")}

    ar_turn = safe_div(rev_now, ar)
    ar_turn_5y = safe_div(rev_5y, ar_5y)
    ar_chg = ar_turn - ar_turn_5y if (ar_turn is not None and ar_turn_5y is not None) else None
    tree["売上債権回転率"] = {"現在値": ar_turn, "5年変化": ar_chg, "評価": "改善" if ar_chg and ar_chg > 0 else ("悪化" if ar_chg and ar_chg < 0 else "横ばい")}

    inv_turn = safe_div(cogs, inventory) if cogs else safe_div(rev_now, inventory)
    inv_turn_5y = safe_div(cogs_5y, inventory_5y) if cogs_5y else safe_div(rev_5y, inventory_5y)
    inv_chg = inv_turn - inv_turn_5y if (inv_turn is not None and inv_turn_5y is not None) else None
    tree["棚卸資産回転率"] = {"現在値": inv_turn, "5年変化": inv_chg, "評価": "改善" if inv_chg and inv_chg > 0 else ("悪化" if inv_chg and inv_chg < 0 else "横ばい")}

    # Sub-decomposition of Equity Multiplier
    equity_ratio = d.get("equity_ratio")
    debt_ratio = 100 - equity_ratio if equity_ratio is not None else None
    tree["自己資本比率"] = {"現在値": equity_ratio}
    tree["負債比率"] = {"現在値": debt_ratio}

    # ROE Contribution ranking (DuPont factors)
    roe_contributions = {}
    if npm_chg is not None and at is not None and em is not None:
        roe_contributions["純利益率"] = npm_chg * 0.01 * (at_5y or at) * (em_5y or em) * 100
    if at_chg is not None and npm is not None and em is not None:
        roe_contributions["総資産回転率"] = at_chg * (npm_5y or npm) * (em_5y or em) * 100
    if em_chg is not None and npm is not None and at is not None:
        roe_contributions["財務レバレッジ"] = em_chg * (npm_5y or npm) * (at_5y or at) * 100

    ranked = sorted(
        [(k, v) for k, v in roe_contributions.items() if v is not None],
        key=lambda x: abs(x[1]), reverse=True
    )
    tree["貢献度ランキング"] = [{"順位": i + 1, "指標": k, "改善寄与度": v, "評価": "改善" if v > 0 else "悪化"} for i, (k, v) in enumerate(ranked)]

    return tree


def compute_pbr_contribution(roe_tree, screening, data, benchmark=None):
    """PBR = ROE × PER / 100 の変化要因を分解

    PBR変化 = ΔROE × PER_5y / 100 + ROE_5y × ΔPER / 100 + ΔROE × ΔPER / 100
    ROE要因はDuPont 3因子に分解し、PER変動も独立要因として加える
    """
    th = generate_dynamic_thresholds(benchmark)
    per_data = screening.get("C-1_PER", {})
    pbr_data = screening.get("C-2_PBR", {})
    per_now = per_data.get("実績値")
    pbr_now = pbr_data.get("実績値")

    per_5y = data.get("per_5y")
    pbr_5y = data.get("pbr_5y")

    per_chg = per_now - per_5y if (per_now is not None and per_5y is not None) else None
    pbr_chg = pbr_now - pbr_5y if (pbr_now is not None and pbr_5y is not None) else None

    roe_info = roe_tree.get("ROE", {})
    roe_now = roe_info.get("現在値")
    roe_chg = roe_info.get("5年変化pt")

    # PBR寄与度の計算
    # PBR = ROE × PER / 100
    # ΔPBR ≈ ΔROE_npm × PER_5y/100 + ΔROE_at × PER_5y/100 + ΔROE_lev × PER_5y/100 + ROE_5y × ΔPER/100 + 交差項
    pbr_factors = {}
    roe_ranking = roe_tree.get("貢献度ランキング", [])
    roe_5y_val = roe_now - roe_chg if (roe_now is not None and roe_chg is not None) else None

    # ROE各因子のPBRへの寄与 = ΔROE_factor × PER_5y / 100
    if per_5y is not None:
        for item in roe_ranking:
            name = item["指標"]
            roe_contrib = item["改善寄与度"]
            pbr_impact = roe_contrib * per_5y / 100
            pbr_factors[name] = {
                "pbr_impact": round(pbr_impact, 4),
                "roe_contrib": round(roe_contrib, 2),
                "category": "ROE要因",
            }

    # PER変動のPBRへの寄与 = ROE_5y × ΔPER / 100 + ΔROE × ΔPER / 100 (交差項込み)
    if per_chg is not None and roe_5y_val is not None:
        per_pbr_impact = roe_5y_val * per_chg / 100
        # 交差項もPER変動に含める
        if roe_chg is not None:
            per_pbr_impact += roe_chg * per_chg / 100
        pbr_factors["PER変動"] = {
            "pbr_impact": round(per_pbr_impact, 4),
            "roe_contrib": None,
            "category": "市場評価要因",
        }

    # 上昇要因ランキング（PBR寄与度 > 0、大きい順）
    up_ranked = sorted(
        [(k, v) for k, v in pbr_factors.items() if v["pbr_impact"] > 0],
        key=lambda x: x[1]["pbr_impact"], reverse=True,
    )
    # 下落要因ランキング（PBR寄与度 < 0、絶対値が大きい順）
    down_ranked = sorted(
        [(k, v) for k, v in pbr_factors.items() if v["pbr_impact"] < 0],
        key=lambda x: x[1]["pbr_impact"],
    )

    def build_list(ranked_items):
        result = []
        for i, (name, info) in enumerate(ranked_items):
            result.append({
                "順位": i + 1,
                "指標": name,
                "PBR寄与度": info["pbr_impact"],
                "ROE寄与度pt": info["roe_contrib"],
                "カテゴリ": info["category"],
            })
        return result

    # PBR評価テキスト（動的閾値ベース）
    pbr_lo = th["pbr_lo"]
    pbr_hi = th["pbr_hi"]
    if pbr_now is not None:
        if pbr_now < pbr_lo:
            pbr_eval = "割安"
        elif pbr_now < pbr_lo * 1.5:
            pbr_eval = "やや割安"
        elif pbr_now <= pbr_hi * 0.7:
            pbr_eval = "適正水準"
        elif pbr_now <= pbr_hi:
            pbr_eval = "やや割高"
        else:
            pbr_eval = "割高"
    else:
        pbr_eval = "データなし"

    return {
        "up_ranking": build_list(up_ranked),
        "down_ranking": build_list(down_ranked),
        "pbr_now": pbr_now,
        "pbr_5y": pbr_5y,
        "pbr_change": round(pbr_chg, 4) if pbr_chg is not None else None,
        "per_now": per_now,
        "per_5y": per_5y,
        "per_change": round(per_chg, 2) if per_chg is not None else None,
        "roe_now": roe_now,
        "roe_5y_change": roe_chg,
        "pbr_eval": pbr_eval,
        "pbr_range": f"{pbr_lo}～{pbr_hi}x",
    }


def run_full_analysis(data, benchmark=None):
    q_results = analyze_quantitative(data)
    s_results = analyze_screening(data, q_results, benchmark=benchmark)
    r_results = analyze_roa_tree(data)
    roe_results = analyze_roe_tree(data)
    pbr_contrib = compute_pbr_contribution(roe_results, s_results, data, benchmark=benchmark)
    return {
        "company": data.get("company", "Unknown"),
        "ticker": data.get("ticker", ""),
        "industry": data.get("industry", "製造・サービス"),
        "quantitative": q_results,
        "screening": s_results,
        "roa_tree": r_results,
        "roe_tree": roe_results,
        "pbr_contribution": pbr_contrib,
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
        }
    }
