# 🚀 evaluate_onwafer_thermal (オンウェハ熱特性評価システム)

## 📌 概要（Overview）

本システムは、ウエハステージ（サーマルチャック）上のプローバー測定における温度依存性データを自動解析し、温度変化に伴うSパラメーターの変動、中心周波数シフト、および中心周波数温度係数 (TCF: Temperature Coefficient of Frequency) や挿入損失温度係数を自動算出して可視化・レポート化を行うPythonモジュールです。

広温度範囲（例: -40℃ 〜 +125℃）でのオンウェハ測定データの抽出・解析・比較処理を一括で自動化します。

## 🎯 主な目的（Objective）

* **温度依存性解析の自動化**: 各温度点でのスィープデータ（CSV/S2P）の一括読み込みおよび温度軸に対するパラメーター抽出の完全自動化

* **熱係数の精度向上・自動計算**: 中心周波数の温度係数 (TCF [ppm/℃]) や挿入損失の温度ドリフト係数の回帰分析・自動算出

* **解析結果の標準化**: 温度推移プロット（PNG）および全温度特性を網羅した評価レポート（Excel）の即時出力

## 🛠️ 使用技術（Tech Stack）

* **言語**: Python (3.9+)

* **データ処理・解析**: `pandas`, `numpy`, `scipy` (線形回帰・統計解析)

* **可視化・プロット**: `matplotlib`, `seaborn`

* **レポート生成**: `openpyxl` (Excel出力用)

* **環境・管理**: Git (`.gitignore`, `.gitkeep` によるフォルダ維持)

## ⚙️ システム構成・処理フロー（Workflow）

* **Step 1 (データ読み込み・温度割り当て)**: `inputs/` 内に配置された各温度設定における測定CSV/S2Pファイルを読み込み、温度メタデータの抽出と整合を実行

* **Step 2 (指標計算・プロット生成)**: 各温度での中心周波数 $f_0$、挿入損失、リターンロスを抽出し、最小二乗法による TCF [ppm/℃] の算出と温度別S21重ね描きグラフを生成

* **Step 3 (レポート出力)**: 各温度の算出結果および比較グラフをまとめ、`outputs/` フォルダへ解析レポート（.xlsx / .png）として出力

## 🚀 使い方（Quick Start）

* **💻 1. セットアップ手順**:

  * **1. リポジトリの取得**: `git clone https://github.com/engineer-KOH/Analog_RF_Evaluations.git`

  * **2. 依存ライブラリのインストール**: `pip install pandas numpy matplotlib seaborn openpyxl scipy`

* **👤 2. ツールの実行フロー**:

  * **1. データ配置**: `evaluate_onwafer_thermal/inputs/` フォルダ内に評価対象の温度別測定ファイル（.csv, .s2p等）を配置します。

  * **2. スクリプトの起動**:

    * ターミナルで `evaluate_onwafer_thermal/src/` へ移動し、以下を実行します。

      `python evaluate_onwafer_thermal.py`

  * **3. 成果物の確認**: 処理が完了すると、`evaluate_onwafer_thermal/outputs/` フォルダ内に解析結果レポート（Excel）および熱特性グラフ（PNG）が出力されます。

## 📁 フォルダ構成（Repository Structure）

```
evaluate_onwafer_thermal/
├── README.md                                    # モジュール解説ドキュメント
├── .gitignore                                   # モジュール用除外設定
├── inputs/                                      # 評価対象の測定データ（CSV/S2P等）の格納先
├── outputs/                                     # 自動生成された解析レポート・グラフの出力先
└── src/                                         # 評価処理プログラム
    └── evaluate_onwafer_thermal.py              # 熱特性評価メインパイプライン
```