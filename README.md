# Analog_RF_Evaluations (RF Semiconductor Evaluation Suite)

半導体RFフィルターおよび高周波アナログデバイスにおける測定データ処理・波形解析・グラフ化を自動化・標準化するためのPythonスクリプト・ライブラリ集です。
本リポジトリでは、評価カテゴリごとに独立したモジュール構成を採り、入力データ（`inputs`）と評価レポート・出力グラフ（`outputs`）の分離管理を実現しています。

---

## 📂 リポジトリ構成 (Directory Structure)

本リポジトリは以下の5つの評価モジュールで構成されています。各モジュール配下は `src`（実装コード）、`inputs`（入力測定データ）、`outputs`（評価結果・グラフ）の統一したディレクトリ構造を持っています。

```text
Analog_RF_Evaluations/
├── README.md                              # 本ドキュメント
├── .gitignore                             # リポジトリ全体共通の除外設定
│
├── evaluate_digital_modulation/           # 1. デジタル変調特性評価
│   ├── inputs/                            # 入力データ（EVN/BER測定ファイル等）
│   ├── outputs/                           # 出力結果（解析レポート等）
│   └── src/                               # 評価・処理スクリプト (Python)
│
├── evaluate_distortion/                   # 2. RF大信号・歪み特性評価
│   ├── inputs/                            # 入力データ（IP3/P1dB/高調波測定ファイル等）
│   ├── outputs/                           # 出力結果（歪み特性グラフ等）
│   └── src/                               # 評価・処理スクリプト (Python)
│
├── evaluate_linear/                       # 3. 線形特性（Sパラメーター）評価
│   ├── inputs/                            # 入力データ（S2P/Touchstoneファイル等）
│   ├── outputs/                           # 出力結果（Passband/Stopbandグラフ等）
│   └── src/                               # 評価・処理スクリプト (Python)
│
├── evaluate_onwafer_thermal/              # 4. オンウェハ熱特性評価
│   ├── inputs/                            # 入力データ（温度変動測定データ等）
│   ├── outputs/                           # 出力結果（熱特性推移グラフ等）
│   └── src/                               # 評価・処理スクリプト (Python)
│
└── rf_noise_characteristics_evaluation/   # 5. RFノイズ特性（NF）評価
    ├── inputs/                            # 入力データ（Noise Figure測定データ等）
    ├── outputs/                           # 出力結果（NF vs Frequencyグラフ等）
    └── src/                               # 評価・処理スクリプト (Python)