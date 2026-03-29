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
from analyzer import run_full_analysis, INDUSTRY_LIST
from excel_parser import parse_excel, scan_available_metrics, extract_custom_timeseries

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static'),
    instance_path=os.path.join(BASE_DIR, 'instance'),
)
app.config['MAX_CONTENT_LENGTH'] = 80 * 1024 * 1024  # 80MB (5 Excel files)
app.secret_key = os.environ.get('SECRET_KEY', 'fin-analysis-secret-key')

# 一時Excelファイルのパスを保持
_uploaded_excel_path = None

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
                    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                        f.save(tmp.name)
                        try:
                            data, ts_data = parse_excel(tmp.name)
                        finally:
                            os.unlink(tmp.name)
                    # フォームから追加情報を取得
                    company = request.form.get('company', '')
                    ticker = request.form.get('ticker', '')
                    industry = request.form.get('industry', '製造・サービス')
                    if company:
                        data['company'] = company
                    if ticker:
                        data['ticker'] = ticker
                    data['industry'] = industry
                    # 定性評価
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

        # Damodaranベンチマークを取得（フロントから業種名が送られてくる場合）
        selected_industry = request.form.get('damodaran_industry', '')
        benchmark = _damodaran_data.get(selected_industry)

        result = run_full_analysis(data, benchmark=benchmark)
        if ts_data:
            result['timeseries'] = ts_data
        # 動的閾値もレスポンスに含める（フロント側で表示用）
        if benchmark:
            from analyzer import generate_dynamic_thresholds
            result['dynamic_thresholds'] = generate_dynamic_thresholds(benchmark)
        return jsonify(result)
    except ImportError as e:
        return jsonify({'error': f'必要なライブラリが不足しています: {str(e)}'}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
        err_msg = str(e)
        if 'openpyxl does not support' in err_msg or '.xls' in err_msg:
            return jsonify({'error': '.xlsファイルの読み込みに失敗しました。xlrdライブラリをインストールしてください。'}), 500
        return jsonify({'error': f'分析中にエラーが発生しました: {err_msg}'}), 500


@app.route('/api/sample')
def sample():
    data = load_sample_data()
    # サンプルのExcelも解析して時系列データを取得
    excel_path = os.path.join(DATA_DIR, '6269-financials.xlsx')
    ts_data = None
    if os.path.exists(excel_path):
        _, ts_data = parse_excel(excel_path)
    # デモではデフォルトの業界を使用
    selected_industry = request.args.get('damodaran_industry', '')
    benchmark = _damodaran_data.get(selected_industry)
    result = run_full_analysis(data, benchmark=benchmark)
    if ts_data:
        result['timeseries'] = ts_data
    return jsonify(result)


@app.route('/api/competitor_analyze', methods=['POST'])
def competitor_analyze():
    """最大5社のExcelファイルを受け取り、比較分析データを返す"""
    try:
        files = request.files.getlist('files[]')
        names = request.form.getlist('names[]')
        industry = request.form.get('damodaran_industry', '')
        benchmark = _damodaran_data.get(industry)

        if not files or len(files) > 5:
            return jsonify({'error': '1〜5社のファイルをアップロードしてください'}), 400

        companies = []
        for i, f in enumerate(files):
            if not f.filename:
                continue
            ext = os.path.splitext(f.filename)[1].lower()
            if ext not in ('.xlsx', '.xls'):
                continue
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                f.save(tmp.name)
                try:
                    data, ts_data = parse_excel(tmp.name)
                finally:
                    os.unlink(tmp.name)

            name = names[i] if i < len(names) and names[i].strip() else os.path.splitext(f.filename)[0]
            data['industry'] = industry or ''
            result = run_full_analysis(data, benchmark=benchmark)

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
            from analyzer import generate_dynamic_thresholds
            resp['dynamic_thresholds'] = generate_dynamic_thresholds(benchmark)
        return jsonify(resp)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'競合分析エラー: {str(e)}'}), 500


@app.route('/api/scan_metrics', methods=['POST'])
def scan_metrics():
    """アップロードされたExcelから可視化可能なメトリクスをスキャン"""
    global _uploaded_excel_path
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルが提供されていません'}), 400
    f = request.files['file']
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.xlsx', '.xls'):
        return jsonify({'error': 'Excelファイルのみ対応'}), 400
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
    f.save(tmp.name)
    tmp.close()
    _uploaded_excel_path = tmp.name
    try:
        metrics = scan_available_metrics(tmp.name)
        return jsonify({'metrics': metrics})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/custom_analysis', methods=['POST'])
def custom_analysis():
    """選択されたメトリクスのカスタム分析データを返す"""
    global _uploaded_excel_path
    body = request.get_json()
    selected = body.get('selected', [])
    if not selected:
        return jsonify({'error': '指標が選択されていません'}), 400
    excel_path = _uploaded_excel_path
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
