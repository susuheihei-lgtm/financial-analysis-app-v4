"""
個別株式分析ウェブアプリケーション
Flask ベースのダッシュボード
"""
import os
import sys
import json
import tempfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

from flask import Flask, render_template, request, jsonify, session
from analyzer import run_full_analysis, INDUSTRY_LIST, generate_dynamic_thresholds
from excel_parser import parse_excel, scan_available_metrics, extract_custom_timeseries
from yfinance_parser import parse_yfinance

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static'),
    instance_path=os.path.join(BASE_DIR, 'instance'),
)
app.config['MAX_CONTENT_LENGTH'] = 80 * 1024 * 1024  # 80MB (5 Excel files)
app.secret_key = os.environ.get('SECRET_KEY', 'fin-analysis-secret-key')

DATA_DIR = os.path.join(BASE_DIR, 'data')

# Load Damodaran industry data at startup
_damodaran_data = {}
_damodaran_path = os.path.join(DATA_DIR, 'damodaran_industry.json')
if os.path.exists(_damodaran_path):
    with open(_damodaran_path, 'r', encoding='utf-8') as _f:
        _damodaran_raw = json.load(_f)
        _damodaran_data = _damodaran_raw.get('industries', {})


def load_sample_data():
    path = os.path.join(DATA_DIR, 'stock_data.json')
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def _build_analysis_response(data, ts_data, benchmark, investor_profile):
    """共通の分析実行・レスポンス構築ヘルパー。
    全分析エンドポイントで同一ロジックを使うことで、追加・変更の漏れを防ぐ。
    """
    result = run_full_analysis(data, benchmark=benchmark, investor_profile=investor_profile)
    if ts_data:
        result['timeseries'] = ts_data
    if benchmark:
        result['dynamic_thresholds'] = generate_dynamic_thresholds(benchmark, profile=investor_profile)
    return result


@app.route('/')
def index():
    return render_template('index.html', industries=INDUSTRY_LIST)


@app.route('/api/damodaran_industries')
def damodaran_industries():
    """Return list of industry names from Damodaran data."""
    names = sorted(_damodaran_data.keys())
    return jsonify(names)


@app.route('/api/industry_benchmark')
def industry_benchmark():
    """Return benchmark data for a given industry."""
    industry = request.args.get('industry', '')
    if industry in _damodaran_data:
        return jsonify(_damodaran_data[industry])
    return jsonify({'error': 'Industry not found'}), 404


@app.route('/api/analyze', methods=['POST'])
def analyze():
    ts_data = None

    try:
        if 'file' in request.files:
            f = request.files['file']
            if f.filename:
                ext = os.path.splitext(f.filename)[1].lower()
                if ext in ('.xlsx', '.xls'):
                    currency = request.form.get('currency', 'JPY')
                    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                        f.save(tmp.name)
                        try:
                            data, ts_data = parse_excel(tmp.name, currency=currency)
                        finally:
                            os.unlink(tmp.name)
                    # フォームから追加情報を取得
                    company = request.form.get('company', '').strip()
                    ticker = request.form.get('ticker', '').strip()
                    industry = request.form.get('industry', '製造・サービス')

                    # Set company/ticker from form, or fallback to filename
                    if company:
                        data['company'] = company
                    elif not data.get('company'):
                        data['company'] = os.path.splitext(f.filename)[0]

                    if ticker:
                        data['ticker'] = ticker
                    elif not data.get('ticker'):
                        import re
                        match = re.search(r'^(\d{4,5})|[-_](\d{4,5})[-_]', f.filename)
                        if match:
                            data['ticker'] = match.group(1) or match.group(2)

                    data['industry'] = industry
                    data['d1_mgmt_change'] = request.form.get('d1', '○')
                    data['d2_ownership'] = request.form.get('d2', '○')
                    data['d3_esg'] = request.form.get('d3', '○')
                elif ext == '.json':
                    data = json.load(f)
                else:
                    return jsonify({'error': 'サポートされていないファイル形式です'}), 400
            else:
                return jsonify({'error': 'ファイルが選択されていません'}), 400
        elif request.is_json:
            data = request.get_json()
        else:
            return jsonify({'error': 'データが提供されていません'}), 400

        selected_industry = request.form.get('damodaran_industry', '')
        benchmark = _damodaran_data.get(selected_industry)
        investor_profile = request.form.get('investor_profile', 'balanced')

        return jsonify(_build_analysis_response(data, ts_data, benchmark, investor_profile))
    except ImportError as e:
        return jsonify({'error': f'必要なライブラリが不足しています: {str(e)}'}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
        err_msg = str(e)
        if 'openpyxl does not support' in err_msg or '.xls' in err_msg:
            return jsonify({'error': '.xlsファイルの読み込みに失敗しました。xlrdライブラリをインストールしてください。'}), 500
        return jsonify({'error': f'分析中にエラーが発生しました: {err_msg}'}), 500


@app.route('/api/fetch_ticker', methods=['POST'])
def fetch_ticker():
    """ティッカーシンボルからyfinanceでデータを取得して分析"""
    try:
        body = request.get_json() or {}
        symbol = body.get('ticker', '').strip()
        if not symbol:
            return jsonify({'error': 'ティッカーシンボルを入力してください'}), 400

        data, ts_data = parse_yfinance(symbol)

        industry = body.get('industry', '')
        if industry:
            data['industry'] = industry
        data['d1_mgmt_change'] = body.get('d1', '○')
        data['d2_ownership'] = body.get('d2', '○')
        data['d3_esg'] = body.get('d3', '○')

        damodaran_industry = body.get('damodaran_industry', '')
        benchmark = _damodaran_data.get(damodaran_industry)
        investor_profile = body.get('investor_profile', 'balanced')

        return jsonify(_build_analysis_response(data, ts_data, benchmark, investor_profile))
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'データ取得中にエラーが発生しました: {str(e)}'}), 500


@app.route('/api/sample')
def sample():
    data = load_sample_data()
    excel_path = os.path.join(DATA_DIR, '6269-financials.xlsx')
    ts_data = None
    if os.path.exists(excel_path):
        _, ts_data = parse_excel(excel_path)
    selected_industry = request.args.get('damodaran_industry', '')
    benchmark = _damodaran_data.get(selected_industry)
    investor_profile = request.args.get('investor_profile', 'balanced')
    return jsonify(_build_analysis_response(data, ts_data, benchmark, investor_profile))


@app.route('/api/competitor_analyze', methods=['POST'])
def competitor_analyze():
    """最大5社のExcel/ティッカーを受け取り、比較分析データを返す"""
    try:
        slot_types = request.form.getlist('types[]')
        names = request.form.getlist('names[]')
        tickers = request.form.getlist('tickers[]')
        files = request.files.getlist('files[]')
        industry = request.form.get('damodaran_industry', '')
        benchmark = _damodaran_data.get(industry)
        investor_profile = request.form.get('investor_profile', 'balanced')

        companies = []
        file_idx = 0
        ticker_idx = 0

        for i, slot_type in enumerate(slot_types):
            name = names[i] if i < len(names) and names[i].strip() else ''

            if slot_type == 'ticker':
                symbol = tickers[ticker_idx] if ticker_idx < len(tickers) else ''
                ticker_idx += 1
                if not symbol.strip():
                    continue
                data, ts_data = parse_yfinance(symbol.strip())
                if not name:
                    name = data.get('company', symbol)
            else:
                if file_idx >= len(files):
                    continue
                f = files[file_idx]
                file_idx += 1
                if not f.filename:
                    continue
                ext = os.path.splitext(f.filename)[1].lower()
                if ext not in ('.xlsx', '.xls'):
                    continue
                comp_currency = request.form.get('currency', 'JPY')
                with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                    f.save(tmp.name)
                    try:
                        data, ts_data = parse_excel(tmp.name, currency=comp_currency)
                    finally:
                        os.unlink(tmp.name)
                if not name:
                    name = os.path.splitext(f.filename)[0]

            data['industry'] = industry or ''
            result = run_full_analysis(data, benchmark=benchmark, investor_profile=investor_profile)

            companies.append({
                'name': name,
                'timeseries': ts_data,
                'screening': result.get('screening', {}),
                'quantitative': result.get('quantitative', {}),
            })

        if len(companies) < 2:
            return jsonify({'error': '比較には2社以上が必要です'}), 400

        resp = {'companies': companies}
        if benchmark:
            resp['dynamic_thresholds'] = generate_dynamic_thresholds(benchmark, profile=investor_profile)
        return jsonify(resp)
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'競合分析エラー: {str(e)}'}), 500


@app.route('/api/scan_metrics', methods=['POST'])
def scan_metrics():
    """アップロードされたExcelから可視化可能なメトリクスをスキャン"""
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルが提供されていません'}), 400
    f = request.files['file']
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.xlsx', '.xls'):
        return jsonify({'error': 'Excelファイルのみ対応'}), 400
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
    f.save(tmp.name)
    tmp.close()
    # セッションにパスを保存（グローバル変数によるレースコンディション回避）
    session['uploaded_excel_path'] = tmp.name
    try:
        metrics = scan_available_metrics(tmp.name)
        return jsonify({'metrics': metrics})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/custom_analysis', methods=['POST'])
def custom_analysis():
    """選択されたメトリクスのカスタム分析データを返す"""
    body = request.get_json()
    selected = body.get('selected', [])
    if not selected:
        return jsonify({'error': '指標が選択されていません'}), 400
    # セッションからパスを取得（各ユーザーのアップロードを独立して管理）
    excel_path = session.get('uploaded_excel_path')
    if not excel_path or not os.path.exists(excel_path):
        excel_path = os.path.join(DATA_DIR, '6269-financials.xlsx')
        if not os.path.exists(excel_path):
            return jsonify({'error': 'Excelファイルが見つかりません'}), 400
    try:
        ts = extract_custom_timeseries(excel_path, selected)
        return jsonify({'timeseries': ts, 'selected': selected})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/scan_sample')
def scan_sample():
    """サンプルExcelのメトリクスをスキャン"""
    excel_path = os.path.join(DATA_DIR, '6269-financials.xlsx')
    if not os.path.exists(excel_path):
        return jsonify({'metrics': []})
    try:
        metrics = scan_available_metrics(excel_path)
        return jsonify({'metrics': metrics})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    os.chdir(BASE_DIR)
    app.run(debug=True, port=int(os.environ.get("PORT", 5050)), load_dotenv=False)
