# 🚀 rf_noise_characteristics_evaluation (RFノイズ特性評価システム)

## 📌 概要（Overview）

本システムは、LNA（低ノイズアンプ）やRFフィルター、各種高周波受動素子におけるノイズ特性を自動解析し、ノイズフィギュア (NF: Noise Figure)、ノイズ温度 (Equivalent Noise Temperature)、および最少ノイズフィギュア ($NF_{\text{min}}$) や最適反射係数 ($\Gamma_{\text{opt}}$) などの評価指標の抽出からプロット・レポート生成を一括して行うPythonモジュールです。

ノイズフィギュアアナライザやスペクトラムアナライザのYファクタ法測定ログ（CSV）を取り込み、周波数依存性のグラフ化および自動集計を完全自動化します。

## 🎯 主な目的（Objective）

* **ノイズ特性解析の自動化**: Yファクタ法データやNF測定ログ（CSV）の自動読み込み・ゲイン補正・ノイズパラメータ計算のパイプライン化

* **主要ノイズ指標の一括可視化**: 周波数に対する NF \[dB\]、Gain \[dB\]、および等価ノイズ温度 $T_e$ \[K\] の推移のプロット

* **解析結果の標準化**: 高解像度ノイズ特性グラフ（PNG）および評価結果サマリー（Excel）の即時自動生成

## 🛠️ 使用技術（Tech Stack）

* **言語**: Python (3.9+)

* **データ処理・解析**: `pandas`, `numpy`, `scipy`

* **可視化・プロット**: `matplotlib`, `seaborn`

* **レポート生成**: `openpyxl` (Excel出力用)

* **環境・管理**: Git (`.gitignore`, `.gitkeep` によるフォルダ維持)

## ⚙️ システム構成・処理フロー（Workflow）

* **Step 1 (データ読み込み・前処理)**: `inputs/` 内に配置されたNF測定データ（Yファクタログ、ENRテーブル、CSV等）を読み込み、冷/熱ノイズパワーの抽出と整合を実行

* **Step 2 (指標計算・プロット生成)**: 受信機ノイズ補正、ゲインおよび NF \[dB\] の算出、周波数軸に対するノイズフィギュアプロットを自動生成

* **Step 3 (レポート出力)**: 算出パラメータおよび生成されたグラフを統合し、`outputs/` フォルダへ解析レポート（.xlsx / .png）として出力

## 🚀 使い方（Quick Start）

* **💻 1. セットアップ手順**:

  * **1. リポジトリの取得**: `git clone https://github.com/engineer-KOH/Analog_RF_Evaluations.git`

  * **2. 依存ライブラリのインストール**: `pip install pandas numpy matplotlib seaborn openpyxl scipy`

* **👤 2. ツールの実行フロー**:

  * **1. データ配置**: `rf_noise_characteristics_evaluation/inputs/` フォルダ内に評価対象の測定データファイル（.csv等）を配置します。

  * **2. スクリプトの起動**:

    * ターミナルで `rf_noise_characteristics_evaluation/src/` へ移動し、以下を実行します。

      `python rf_noise_characteristics_evaluation.py`

  * **3. 成果物の確認**: 処理が完了すると、`rf_noise_characteristics_evaluation/outputs/` フォルダ内に解析結果レポート（Excel）およびノイズ特性グラフ（PNG）が出力されます。

## 📁 フォルダ構成（Repository Structure）

```
rf_noise_characteristics_evaluation/
├── README.md                                           # モジュール解説ドキュメント
├── .gitignore                                          # モジュール用除外設定
├── inputs/                                             # 評価対象の測定データ（CSV等）の格納先
├── outputs/                                            # 自動生成された解析レポート・グラフの出力先
└── src/                                                # 評価処理プログラム
    └── rf_noise_characteristics_evaluation.py          # ノイズ特性評価メインパイプライン

```