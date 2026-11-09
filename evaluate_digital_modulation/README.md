# 🚀 evaluate_digital_modulation (デジタル変調特性評価システム)

## 📌 概要（Overview）
本システムは、QAMやOFDMなどのデジタル変調信号における測定ログおよびIQ波形データを自動解析し、EVM (Error Vector Magnitude)、コンスタレーション（Constellation）、BER (Bit Error Rate) などの評価指標の解析からグラフ化・レポート自動生成までを一括して行うPythonモジュールです。  
高周波変調評価の手作業によるデータ抽出・グラフ作成工数を大幅に削減し、標準化された評価レポートを出力します。

## 🎯 主な目的（Objective）
- **データ解析・可視化の自動化**: 測定ログ（CSV）の読み込みからEVM算出、コンスタレーション描画、結果出力までの完全自動パイプライン化
- **品質指標の一括評価**: RMS/Peak EVM、Phase Error、Magnitude Error、BER等のパラメータを単一実行で総合評価
- **解析結果の標準化**: 評価グラフ（PNG）と集計データ（Excel/JSON）を統一フォーマットで即時出力

## 🛠️ 使用技術（Tech Stack）
- **言語**: Python (3.9+)
- **データ処理・解析**: `pandas`, `numpy`, `scipy`
- **可視化・プロット**: `matplotlib`, `seaborn`
- **レポート生成**: `openpyxl` (Excel出力用)
- **環境・管理**: Git (`.gitignore`, `.gitkeep` によるフォルダ維持)

## ⚙️ システム構成・処理フロー（Workflow）
- **Step 1 (データ読み込み・前処理)**: `inputs/` 内に配置された測定CSVファイル（IQデータ、EVM測定ログ等）を読み込み、不要なノイズやヘッダーの除外・データ整形を実行
- **Step 2 (指標計算・プロット生成)**: 整形済みデータからEVM[%]（RMS/Peak）、シンボル点配置（Constellation）、IQミスマッチ等を算出し、高解像度グラフを生成
- **Step 3 (レポート出力)**: 算出結果および生成されたグラフ群を統合し、`outputs/` フォルダへ解析レポート（.xlsx / .png）として出力

## 🚀 使い方（Quick Start）
- **💻 1. セットアップ手順**:
  - **1. リポジトリの取得**: `git clone https://github.com/engineer-KOH/Analog_RF_Evaluations.git`
  - **2. 依存ライブラリのインストール**: `pip install pandas numpy matplotlib seaborn openpyxl scipy`

- **👤 2. ツールの実行フロー**:
  - **1. データ配置**: `evaluate_digital_modulation/inputs/` フォルダ内に評価対象の測定ログファイル（.csv等）を配置します。
  - **2. スクリプトの起動**:
    - ターミナルで `evaluate_digital_modulation/src/` へ移動し、以下を実行します。  
      `python evaluate_digital_modulation.py`
  - **3. 成果物の確認**: 処理が完了すると、`evaluate_digital_modulation/outputs/` フォルダ内に解析結果レポート（Excel）およびグラフ（PNG）が出力されます。

## 📁 フォルダ構成（Repository Structure）
```text
evaluate_digital_modulation/
├── README.md                                    # モジュール解説ドキュメント
├── .gitignore                                   # モジュール用除外設定
├── inputs/                                      # 評価対象の測定データ（CSV等）の格納先
├── outputs/                                     # 自動生成された解析レポート・グラフの出力先
└── src/                                         # 評価処理プログラム
    └── evaluate_digital_modulation.py           # 変調特性評価メインパイプライン
```