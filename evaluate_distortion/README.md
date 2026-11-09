# 🚀 evaluate_distortion (RF歪み・大信号特性評価システム)

## 📌 概要（Overview）

本システムは、RFパワーアンプやRFフィルター等の高周波デバイスにおける歪み特性および大信号入出力特性を自動解析し、1dB利得抑圧点 ($P_{\text{1dB}}$)、3次高調波歪み ($IMD_3$)、3次相互変調歪み ($IP_3 / IMD_3$) などの評価指標の抽出から可視化・レポート生成までを一括して行うPythonモジュールです。

歪み評価に伴う煩雑な利得抑圧点や交点（Intercept Point）の算出処理を完全自動化し、再現性の高い評価データを提供します。

## 🎯 主な目的（Objective）

* **大信号・非線形解析の自動化**: 測定ログ（電力スィープCSV等）から利得抑圧曲線や歪み交点を自動フィッティング・算出

* **主要歪み指標の一括抽出**: $P_{\text{1dB}}$（Input/Output）、$IIP_3 / OIP_3$、高調波レベル（2nd, 3rd harmonics）の自動算出

* **解析結果の標準化**: 電力依存性プロット（PNG）および評価結果サマリー（Excel/JSON）を統一フォーマットで自動生成

## 🛠️ 使用技術（Tech Stack）

* **言語**: Python (3.9+)

* **データ処理・解析**: `pandas`, `numpy`, `scipy` (多項式フィッティング・交点計算)

* **可視化・プロット**: `matplotlib`, `seaborn`

* **レポート生成**: `openpyxl` (Excel出力用)

* **環境・管理**: Git (`.gitignore`, `.gitkeep` によるフォルダ維持)

## ⚙️ システム構成・処理フロー（Workflow）

* **Step 1 (データ読み込み・前処理)**: `inputs/` 内に配置されたスィープ測定CSVファイル（Pin vs Pout, F1/F2 2トーンスペクトラムログ等）を読み込み、データ補正および構造化を実行

* **Step 2 (指標計算・プロット生成)**: 線形領域の最小二乗フィッティングによる $P_{\text{1dB}}$ 検出、および 1st / 3rd オーダー線の交点計算による $IP_3$ 抽出を行い、入出力電力特性グラフを生成

* **Step 3 (レポート出力)**: 算出されたパラメータ群および生成グラフをまとめ、`outputs/` フォルダへ解析レポート（.xlsx / .png）として出力

## 🚀 使い方（Quick Start）

* **💻 1. セットアップ手順**:

  * **1. リポジトリの取得**: `git clone https://github.com/engineer-KOH/Analog_RF_Evaluations.git`

  * **2. 依存ライブラリのインストール**: `pip install pandas numpy matplotlib seaborn openpyxl scipy`

* **👤 2. ツールの実行フロー**:

  * **1. データ配置**: `evaluate_distortion/inputs/` フォルダ内に評価対象の測定データファイル（.csv等）を配置します。

  * **2. スクリプトの起動**:

    * ターミナルで `evaluate_distortion/src/` へ移動し、以下を実行します。

      `python evaluate_distortion.py`

  * **3. 成果物の確認**: 処理が完了すると、`evaluate_distortion/outputs/` フォルダ内に解析結果レポート（Excel）および歪み特性グラフ（PNG）が出力されます。

## 📁 フォルダ構成（Repository Structure）

```
evaluate_distortion/
├── README.md                                    # モジュール解説ドキュメント
├── .gitignore                                   # モジュール用除外設定
├── inputs/                                      # 評価対象の測定データ（CSV等）の格納先
│   ├── .gitkeep                                 # フォルダ維持用ファイル
│   └── sample_distortion_measurement_data.csv   # 入力サンプルデータ
├── outputs/                                     # 自動生成された解析レポート・グラフの出力先
│   ├── .gitkeep                                 # フォルダ維持用ファイル
│   └── rf_distortion_evaluation_report.xlsx     # 自動生成レポート例
└── src/                                         # 評価処理プログラム
    └── evaluate_distortion.py                   # 歪み特性評価メインパイプライン
```