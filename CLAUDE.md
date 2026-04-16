# Financial Analysis Automation v5

Flask製の個別株式分析ウェブアプリ。ティッカー入力またはExcelアップロードで財務データを取得し、スコアリング・チャート・競合比較を1画面で提供。

## Commands

| Command | Description |
|---------|-------------|
| `.venv/bin/python3 app.py` | 開発サーバー起動（port 5050） |
| `.venv/bin/python3 -m pip install -r requirements.txt` | 依存パッケージインストール |
| `lsof -i :5050 -t \| xargs kill -9` | ポート5050を強制解放 |
| `git push origin main` | GitHub v5リポジトリへpush |

## Architecture

```
v5/
  app.py                    # Flaskルート定義（API + セッション管理）
  analyzer.py               # 分析エンジン（薄いオーケストレーター）
  _analyzer_helpers.py      # 共通ユーティリティ（safe_div, rate_change等）
  _analyzer_quantitative.py # 定量スコアリングロジック
  _analyzer_screening.py    # スクリーニング判定
  _analyzer_thresholds.py   # 投資家プロファイル・しきい値定義
  _analyzer_trees.py        # ROA/ROEツリー分解
  yfinance_parser.py        # yfinance + SEC EDGAR データ取得
  excel_parser.py           # Excelファイルパーサー（.xls/.xlsx、日本語対応）
  templates/index.html      # シングルページアプリ（HTML/CSS/JS 3153行）
  static/css/, static/js/   # 静的アセット
  data/                     # サンプルExcelデータ
  .env                      # ローカル環境変数（gitignore済み）
```

## Key Files

- `app.py:178` — `/api/analyze` メインエンドポイント（POST, Excel or tickerデータ受け取り）
- `app.py:240` — `/api/fetch_ticker` yfinance/EDGAR でティッカー情報取得
- `yfinance_parser.py:645` — `parse_yfinance()` エントリポイント
- `yfinance_parser.py:218` — `_get_sec_annual_series()` SEC EDGAR年次データ取得
- `_analyzer_thresholds.py` — 投資家プロファイル5種（conservative/balanced/growth/income/aggressive）
- `templates/index.html` — V0エディトリアルデザイン（Bebas Neue + IBM Plex Mono + #D4852A）

## Environment

```bash
# .env（必須 — ないとSECRET_KEY RuntimeErrorで起動不可）
FLASK_DEBUG=true
SECRET_KEY=<32バイトhex>  # python -c "import secrets; print(secrets.token_hex(32))"
```

- `python-dotenv` が venv にインストールされている必要あり（`pip install python-dotenv`）
- `PORT` 環境変数でポート変更可能（デフォルト: 5050）

## Gotchas

- **yfinance は年次データ4年まで**の仕様制限。5年目は取得不可（米国株はSEC EDGARで5年取得可）
- **SEC EDGARは米国株のみ**。日本株（`.T`）や南アフリカ株（AU等）はyfinanceのみ
- **AU（AngloGold）はSEC EDGARにus-gaapデータなし** — yfinanceフォールバックで処理される
- **`major_holders` フォーマット変更（yfinance新版）**: string index + 'Value'列。旧コードの`iloc[str, 1]`はクラッシュする — `_assess_ownership()`で対応済み
- **起動前にObsidianを開く必要あり** — MCP Tools（REST API Plugin）がObsidian起動を要求
- **venvのpythonを使うこと** — システムpythonではパッケージが見つからない
- **テンプレートは`templates/index.html`の1ファイルのみ** — CSS/JSもすべてインライン

## Design System (V0 Editorial)

- フォント: `Bebas Neue`（見出し）、`IBM Plex Mono`（ラベル/ボタン）
- アクセントカラー: `#D4852A`（CSS変数: `--accent-v0`）
- コーナー: `border-radius: 0`（角丸禁止）
- 背景: グリッドパターン60px + ノイズオーバーレイ

## GitHub

- v5: `https://github.com/susuheihei-lgtm/financial-analysis-app-v5`
- v4: `https://github.com/susuheihei-lgtm/financial-analysis-app-v4`（v5と同内容、Render.comデプロイ用）
