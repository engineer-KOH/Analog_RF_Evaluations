# 🚀 evaluate_linear (RF小信号・線形特性評価システム)

## 📌 概要（Overview）

本システムは、RFフィルターや各種高周波アンプ・受動素子における小信号・線形特性（Sパラメーター）を自動解析し、通過帯域挿入損失 (Passband Insertion Loss)、帯域外減衰量 (Stopband Attenuation)、リターンロス (S11/S22)、Q値 (Quality Factor) などの評価指標の自動抽出・プロット・レポート生成を一括して行うPythonモジュールです。

ネットワークアナライザ等から出力されるTouchstone形式ファイル (.s2p) や各種CSV測定ログを高速解析し、標準化された評価結果を出力します。

## 🎯 主な目的（Objective）

* **Sパラメーター解析の自動化**: Touchstone (.s2p) または CSV形式データの自動読み込み・複素数演算・パラメータ抽出

* **主要線形指標の一括可視化**: 帯域幅 (Bandwidth)、センター周波数 ($f_0$)、Ripple、S21/S11周波数特性の可視化

* **解析結果の標準化**: 高解像度周波数レスポンスプロット（PNG）および評価結果集計（Excel）の自動生成

## 🛠️ 使用技術（Tech Stack）

* **言語**: Python (3.9+)

* **データ処理・解析**: `pandas`, `numpy`, `scikit-rf` (Sパラメーター処理), `scipy`

* **可視化・プロット**: `matplotlib`, `seaborn`

* **レポート生成**: `openpyxl` (Excel出力用)

* **環境・管理**: Git (`.gitignore`, `.gitkeep` によるフォルダ維持)

## ⚙️ システム構成・処理フロー（Workflow）

* **Step 1 (データ読み込み・前処理)**: `inputs/` 内に配置されたSパラメーターファイル (.s2p / .csv) を読み込み、複素数振幅/位相データの抽出および周波数グリッドの整合を実行

* **Step 2 (指標計算・プロット生成)**: 通過帯域 (Passband) の最小挿入損失、帯域端周波数、リターンロスのピーク値、および阻止帯域 (Stopband) 減衰量を自動算出し、S21/S11周波数レスポンスグラフを生成

* **Step 3 (レポート出力)**: 算出されたパラメータ群および生成グラフを統合し、`outputs/` フォルダへ解析レポート（.xlsx / .png）として出力

## 🚀 使い方（Quick Start）

* **💻 1. セットアップ手順**:

  * **1. リポジトリの取得**: `git clone https://github.com/engineer-KOH/Analog_RF_Evaluations.git`

  * **2. 依存ライブラリのインストール**: `pip install pandas numpy matplotlib seaborn openpyxl scikit-rf scipy`

* **👤 2. ツールの実行フロー**:

  * **1. データ配置**: `evaluate_linear/inputs/` フォルダ内に評価対象の測定データファイル（.s2p, .csv等）を配置します。

  * **2. スクリプトの起動**:

    * ターミナルで `evaluate_linear/src/` へ移動し、以下を実行します。

      `python evaluate_linear.py`

  * **3. 成果物の確認**: 処理が完了すると、`evaluate_linear/outputs/` フォルダ内に解析結果レポート（Excel）およびSパラメーター特性グラフ（PNG）が出力されます。

## 📁 フォルダ構成（Repository Structure）

```
evaluate_linear/
├── README.md                                    # モジュール解説ドキュメント
├── .gitignore                                   # モジュール用除外設定
├── inputs/                                      # 評価対象の測定データ（S2P/CSV等）の格納先
│   ├── .gitkeep                                 # フォルダ維持用ファイル
│   └── sample_linear_measurement_data.csv       # 入力サンプルデータ
├── outputs/                                     # 自動生成された解析レポート・グラフの出力先
│   ├── .gitkeep                                 # フォルダ維持用ファイル
│   └── rf_linear_evaluation_report.xlsx         # 自動生成レポート例
└── src/                                         # 評価処理プログラム
    └── evaluate_linear.py                       # 線形特性評価メインパイプライン

```