# 💰 資金繰りリスク可視化ツール

中小企業経営者向けの資金繰りリスク診断ツール

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://cashflow-risk-analyzer-zityxm7xbuivhjmwb5ngvv.streamlit.app/)

---

## ✨ 主な機能

### 💵 Phase 3: 長期運転資金算出（NEW）
- 目標資金ショート確率（5%/10%/15%）から必要資金額を逆算
- 5年返済プラン + 最短償還年数を表示
- 返済可能性を自動判定
- 銀行交渉用レポート自動生成

### 🔥 Phase 2: 歴史的ストレステスト（NEW）
実データに基づくストレステスト：
- リーマンショック（2008年）
- 東日本大震災（2011年）
- コロナショック（2020年）
- ウクライナ戦争（2022年）

### 📄 レポート出力（NEW）
- **Word形式**: 銀行交渉用正式文書
- **Excel形式**: 詳細データ確認用

### 📊 その他の機能
- モンテカルロシミュレーション（10,000回固定）
- 業種別×従業員数別の標準値対応
- サービス業細分化（IT、医療・介護、広告、その他）
- スマホ対応UI

---

## 🚀 使い方

### オンラインで使う（推奨）

**アプリURL**: https://cashflow-risk-analyzer-zityxm7xbuivhjmwb5ngvv.streamlit.app/

1. 上記URLにアクセス
2. 財務数値を入力（8項目のみ）
3. シミュレーション実行ボタンをクリック
4. 結果を確認してレポートをダウンロード

**スマホでも利用可能**: レスポンシブデザイン対応

---

### ローカルで実行

```bash
# リポジトリのクローン
git clone https://github.com/atsushi0820/cashflow-risk-analyzer.git
cd cashflow-risk-analyzer

# 依存ライブラリのインストール
pip install -r requirements.txt

# アプリの実行
streamlit run app.py
```

ブラウザで `http://localhost:8501` が自動的に開きます。

---

## 📊 対応業種

### 基本5業種
- 建設業
- 製造業
- 卸・小売業
- 運輸・物流業
- サービス・その他

### サービス業の細分化（v2.0）
- IT・システム開発
- 医療・介護
- 広告・マーケティング
- その他（士業、コンサル、飲食等）

### 従業員数区分
- 30人以下
- 30人超～50人以下
- 50人超～100人以下
- 100人超～300人以下
- 300人超

**合計40パターン**の業種別標準値を収録

---

## 🎯 このツールの特徴

### 1. 最小限の入力で高品質な分析
わずか8項目の入力で、10,000回のモンテカルロシミュレーションを実行

### 2. 実データベースのストレステスト
恣意的な仮定ではなく、過去の危機の実績データを使用

### 3. 定量的な銀行交渉
「資金ショート確率を5%以下に抑えるには○○万円必要」と定量的に提示

### 4. 銀行実務の知見を反映
20年の金融実務経験（70万社の分析実績）を基盤に設計

---

## 🔧 技術スタック

- **Python**: 3.11
- **Streamlit**: Webアプリフレームワーク
- **NumPy**: モンテカルロシミュレーション
- **Matplotlib**: グラフ描画
- **python-docx**: Word出力
- **openpyxl**: Excel出力
- **Pandas**: データ処理

---

## 📝 データ出典

### 公開統計
- **法人企業統計**（財務省）
- **中小企業実態基本調査**（財務省・中小企業庁）
- **日本政策金融公庫 中小企業経営指標**

### 歴史的ショックデータ
- リーマンショック: 2008-2009年の業種別売上減少率
- 東日本大震災: 2011年のサプライチェーン影響
- コロナショック: 2020年の業種別影響
- ウクライナ戦争: 2022年のエネルギー価格高騰

---

## 📁 ファイル構成

```
cashflow-risk-analyzer/
├── app.py                           # メインアプリ
├── funding_calculator.py            # Phase 3: 長期運転資金算出
├── shock_analyzer.py                # Phase 2: ストレステスト分析
├── shock_monte_carlo.py             # Phase 2: シミュレーター
├── industry_standards_v2.py         # 業種別標準値v2.0
├── industry_standards_matrix_v2.csv # 業種別データv2.0
├── shock_scenarios.csv              # ストレステストデータ
├── ipag.ttf                         # 日本語フォント
├── requirements.txt                 # 依存ライブラリ
└── README.md                        # このファイル
```

---

## 🎨 UI/UX

### スマホ対応
- 2列レイアウト
- 大きなボタン（高さ60px）
- タップしやすいUI

### 視認性向上
- 色付き背景の大きなセクションヘッダー
- Phase 3（緑系）→ Phase 2（赤系）→ レポート（青系）
- カード風の結果表示

---

## 📄 ライセンス

MIT License

---

## 👤 作成者

中小企業の資金繰り支援を目的として開発

**開発背景**:
- 城南信用金庫・東京商工リサーチでの20年の企業分析経験
- 約70万社の分析実績
- 中央官庁の公開報告書作成実績

---

## 📧 お問い合わせ

- GitHub Issues: https://github.com/atsushi0820/cashflow-risk-analyzer/issues
- ツールURL: https://cashflow-risk-analyzer-zityxm7xbuivhjmwb5ngvv.streamlit.app/

---

## 🔄 更新履歴

### v2.0（2026年4月）
- ✨ Phase 3: 長期運転資金算出機能を追加
- ✨ Phase 2: 歴史的ストレステストに全面改訂
- ✨ Word/Excel出力機能を追加
- 🎨 スマホ対応UI、視認性大幅向上
- 📊 業種別標準値v2.0（サービス業細分化）
- 🔧 表示順序変更（Phase 3 → Phase 2 → レポート）

### v1.0（2026年2月）
- 🎉 初版リリース
- Phase 1-5の基本機能
- Streamlit Cloud公開
- 日本語フォント対応
- 利用マニュアルPDF作成

---

**最終更新**: 2026年4月21日 - Streamlit Cloud再デプロイ
