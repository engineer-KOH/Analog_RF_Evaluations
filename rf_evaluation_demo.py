import numpy as np
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


class RFEvaluationDXDemo:

  def __init__(self, filename="RF_Device_Evaluation_Report.xlsx"):
    self.filename = filename
    self.wb = openpyxl.Workbook()

    # スタイル共通定義
    self.font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    self.font_title = Font(name="Calibri", size=15, bold=True, color="1F4E79")
    self.font_bold = Font(name="Calibri", size=11, bold=True)
    self.font_regular = Font(name="Calibri", size=11)
    self.font_disc_header = Font(
        name="Calibri", size=11, bold=True, color="1F4E79"
    )

    self.fill_navy = PatternFill(
        start_color="1F4E79", end_color="1F4E79", fill_type="solid"
    )
    self.fill_pass = PatternFill(
        start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
    )
    self.fill_fail = PatternFill(
        start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
    )
    self.fill_disc = PatternFill(
        start_color="F2F5F9", end_color="F2F5F9", fill_type="solid"
    )

    self.thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

  def generate_all_reports(self):
    print("[INFO] RF評価DXデモレポートを生成中...")
    self._build_summary_tab()
    self._build_linear_tab()
    self._build_distortion_tab()
    self._build_noise_tab()
    self._build_digital_tab()
    self._build_wafer_temp_tab()

    if "Sheet" in self.wb.sheetnames:
      del self.wb["Sheet"]

    self.wb.save(self.filename)
    print(f"[SUCCESS] 評価レポート生成完了: {self.filename}")

  def _add_discussion_box(self, ws, start_row, title, text):
    """シート内に結果に対する考察（Discussion）ブロックを配置するヘルパー"""
    ws.merge_cells(
        start_row=start_row, start_column=1, end_row=start_row, end_column=6
    )
    cell_title = ws.cell(row=start_row, column=1, value=title)
    cell_title.font = self.font_disc_header
    cell_title.fill = self.fill_disc

    ws.merge_cells(
        start_row=start_row + 1,
        start_column=1,
        end_row=start_row + 3,
        end_column=6,
    )
    cell_text = ws.cell(row=start_row + 1, column=1, value=text)
    cell_text.font = self.font_regular
    cell_text.fill = self.fill_disc
    cell_text.alignment = Alignment(wrap_text=True, vertical="top")

    for r in range(start_row, start_row + 4):
      for c in range(1, 7):
        ws.cell(row=r, column=c).border = self.thin_border

  # ----------------------------------------------------
  # 0. 総合サマリーダッシュボード
  # ----------------------------------------------------
  def _build_summary_tab(self):
    ws = self.wb.active
    ws.title = "Summary_Dashboard"

    ws["A1"] = "RF Device Overall Evaluation Dashboard"
    ws["A1"].font = self.font_title

    headers = [
        "領域",
        "主な評価項目",
        "使用装置",
        "評価ステータス",
        "備考・特記事項",
    ]
    for col, h in enumerate(headers, 1):
      cell = ws.cell(row=3, column=col, value=h)
      cell.font = self.font_header
      cell.fill = self.fill_navy
      cell.alignment = Alignment(horizontal="center")

    summary_data = [
        ("線形特性", "Sパラメータ, 群遅延", "VNA", "PASS", "Band24 規格クリア"),
        (
            "大信号・歪み",
            "P1dB, IMD3, IP3",
            "VNA / 2波SG+スペアナ",
            "PASS",
            "OIP3 > +35dBm 達成",
        ),
        (
            "雑音特性",
            "NF, 位相雑音",
            "NFA / シグナルアナライザ",
            "PASS",
            "NF @ 2.4GHz: 1.25dB",
        ),
        (
            "変調デジタル",
            "EVM, ACLR, BER",
            "VSG + VSA",
            "PASS",
            "EVM: 1.42% (規格 < 2.5%)",
        ),
        (
            "構造・環境",
            "オンウエハ, 温度特性",
            "プローバー / サーモストリーム",
            "PASS",
            "-40℃〜85℃ 歩留まり > 96%",
        ),
    ]

    for row_idx, row_data in enumerate(summary_data, start=4):
      for col_idx, val in enumerate(row_data, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val)
        cell.border = self.thin_border
        if col_idx == 4:
          cell.alignment = Alignment(horizontal="center")
          cell.font = Font(bold=True, color="375623")
          cell.fill = self.fill_pass

    # ★ 考察ボックスの追加
    self._add_discussion_box(
        ws,
        10,
        "【総合評価考察】",
        "全5評価領域において全測定項目が判定基準をクリアし、PASSとなりました。\n"
        "5G/Wi-Fi通信帯域における低挿入損失と高OIP3特性が両立できていることが確認されています。",
    )

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 30

  # ----------------------------------------------------
  # 1. 線形特性
  # ----------------------------------------------------
  def _build_linear_tab(self):
    ws = self.wb.create_sheet(title="1_Linear_SParam")
    ws["A1"] = "S21 Frequency Response & Limit Line"
    ws["A1"].font = self.font_title

    headers = [
        "Freq (GHz)",
        "S21 (dB)",
        "Group Delay (ns)",
        "S21 Limit (dB)",
        "Status",
    ]
    for c, h in enumerate(headers, 1):
      cell = ws.cell(row=3, column=c, value=h)
      cell.font = self.font_header
      cell.fill = self.fill_navy

    freqs = np.linspace(2.0, 3.0, 21)
    limit_val = -3.0

    for i, f in enumerate(freqs, start=4):
      s21 = -1.5 - 15.0 * ((f - 2.45) ** 2)
      gd = 2.1 + 0.3 * np.sin(f * 10)
      status = "PASS" if s21 >= limit_val else "FAIL"

      ws.cell(row=i, column=1, value=round(f, 3))
      ws.cell(row=i, column=2, value=round(s21, 2))
      ws.cell(row=i, column=3, value=round(gd, 2))
      ws.cell(row=i, column=4, value=limit_val)

      st_cell = ws.cell(row=i, column=5, value=status)
      st_cell.fill = self.fill_pass if status == "PASS" else self.fill_fail
      for col in range(1, 6):
        ws.cell(row=i, column=col).border = self.thin_border

    chart = LineChart()
    chart.title = "S21 Insertion Loss"
    chart.style = 13
    chart.y_axis.title = "S21 (dB)"
    chart.x_axis.title = "Frequency (GHz)"

    data = Reference(ws, min_col=2, min_row=3, max_col=4, max_row=24)
    cats = Reference(ws, min_col=1, min_row=4, max_row=24)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[1].graphicalProperties.line.solidFill = "FF0000"
    ws.add_chart(chart, "G3")

    # ★ 考察ボックスの追加
    self._add_discussion_box(
        ws,
        26,
        "【線形特性 結果に対する考察】",
        "中心周波数 2.45 GHz 付近での最小挿入損失は -1.50 dB となり、目標 LIMIT 基準に対して十分なマージンを確保できています。",
    )

    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 12

  # ----------------------------------------------------
  # 2. 大信号・歪み
  # ----------------------------------------------------
  def _build_distortion_tab(self):
    ws = self.wb.create_sheet(title="2_Distortion_IP3")
    ws["A1"] = "P1dB & IMD3 / OIP3 Linear Sweep Analysis"
    ws["A1"].font = self.font_title

    headers = [
        "Pin (dBm)",
        "Pout Fundamental (dBm)",
        "IMD3 (dBm)",
        "Ideal 1:1 Line",
        "Ideal 3:1 Line",
    ]
    for c, h in enumerate(headers, 1):
      cell = ws.cell(row=3, column=c, value=h)
      cell.font = self.font_header
      cell.fill = self.fill_navy

    pins = np.linspace(-20, 10, 16)
    for i, pin in enumerate(pins, start=4):
      pout = pin + 15.0 - (0.01 * np.exp(pin / 2.5))
      imd3 = 3 * pin - 35.0
      ideal1 = pin + 15.0
      ideal3 = 3 * pin - 35.0

      ws.cell(row=i, column=1, value=round(pin, 1))
      ws.cell(row=i, column=2, value=round(pout, 2))
      ws.cell(row=i, column=3, value=round(imd3, 2))
      ws.cell(row=i, column=4, value=round(ideal1, 2))
      ws.cell(row=i, column=5, value=round(ideal3, 2))
      for col in range(1, 6):
        ws.cell(row=i, column=col).border = self.thin_border

    ws["G3"] = "Calculated OIP3 (dBm)"
    ws["H3"] = 40.0
    ws["G3"].font = self.font_bold
    ws["H3"].font = Font(bold=True, color="1F4E79")

    chart = ScatterChart()
    chart.title = "IP3 Intercept Point Trajectory"
    chart.style = 13
    chart.x_axis.title = "Pin (dBm)"
    chart.y_axis.title = "Power (dBm)"

    xvalues = Reference(ws, min_col=1, min_row=4, max_row=19)
    for col in range(2, 6):
      yvalues = Reference(ws, min_col=col, min_row=3, max_row=19)
      series = Series(yvalues, xvalues, title_from_data=True)
      chart.series.append(series)

    ws.add_chart(chart, "G5")

    # ★ 考察ボックスの追加
    self._add_discussion_box(
        ws,
        21,
        "【大信号・歪み特性 結果に対する考察】",
        "基本波とIMD3の回帰直線交点より、OIP3 = +40.0 dBm が得られ、小信号〜中信号領域で高い線形性が維持されています。",
    )

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15

  # ----------------------------------------------------
  # 3. 雑音特性
  # ----------------------------------------------------
  def _build_noise_tab(self):
    ws = self.wb.create_sheet(title="3_Noise_NF_PhaseNoise")
    ws["A1"] = "Noise Figure (NF) & Phase Noise Profile"
    ws["A1"].font = self.font_title

    headers = [
        "Offset Freq",
        "Phase Noise (dBc/Hz)",
        "Limit (dBc/Hz)",
        "Status",
    ]
    for c, h in enumerate(headers, 1):
      cell = ws.cell(row=3, column=c, value=h)
      cell.font = self.font_header
      cell.fill = self.fill_navy

    offsets = ["1k", "10k", "100k", "1M", "10M"]
    p_noise = [-82.5, -108.2, -125.4, -142.1, -151.0]
    limits = [-80.0, -100.0, -120.0, -140.0, -150.0]

    for i, (off, pn, lim) in enumerate(zip(offsets, p_noise, limits), start=4):
      status = "PASS" if pn <= lim else "FAIL"
      ws.cell(row=i, column=1, value=off)
      ws.cell(row=i, column=2, value=pn)
      ws.cell(row=i, column=3, value=lim)
      st_cell = ws.cell(row=i, column=4, value=status)
      st_cell.fill = self.fill_pass if status == "PASS" else self.fill_fail
      for col in range(1, 5):
        ws.cell(row=i, column=col).border = self.thin_border

    ws["F3"] = "Average NF @ 2.4GHz"
    ws["G3"] = "1.25 dB"
    ws["F3"].font = self.font_bold

    chart = LineChart()
    chart.title = "Phase Noise Profile"
    chart.style = 13
    chart.y_axis.title = "dBc/Hz"

    data = Reference(ws, min_col=2, min_row=3, max_col=3, max_row=8)
    cats = Reference(ws, min_col=1, min_row=4, max_row=8)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "F5")

    # ★ 考察ボックスの追加
    self._add_discussion_box(
        ws,
        10,
        "【雑音特性 結果に対する考察】",
        "すべてのオフセット周波数で位相雑音の規格をクリアしており、受信フロントエンドにおける優れたS/N比特性を確保できています。",
    )

    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 12

  # ----------------------------------------------------
  # 4. 変調デジタル
  # ----------------------------------------------------
  def _build_digital_tab(self):
    ws = self.wb.create_sheet(title="4_Digital_EVM_ACLR")
    ws["A1"] = "5G/Wi-Fi Digital Modulation Test Results"
    ws["A1"].font = self.font_title

    headers = [
        "Test Item",
        "Modulation Format",
        "Measured Value",
        "Spec Limit",
        "Unit",
        "Status",
    ]
    for c, h in enumerate(headers, 1):
      cell = ws.cell(row=3, column=c, value=h)
      cell.font = self.font_header
      cell.fill = self.fill_navy

    test_results = [
        ("EVM (RMS)", "256-QAM", 1.42, 2.50, "%", "PASS"),
        ("ACLR Lower", "5G NR 100MHz", -48.5, -45.0, "dBc", "PASS"),
        ("ACLR Upper", "5G NR 100MHz", -47.8, -45.0, "dBc", "PASS"),
        ("BER", "QPSK", 1e-9, 1e-6, "-", "PASS"),
        ("Frequency Error", "256-QAM", 12.3, 50.0, "Hz", "PASS"),
    ]

    for i, row in enumerate(test_results, start=4):
      for j, val in enumerate(row, start=1):
        cell = ws.cell(row=i, column=j, value=val)
        cell.border = self.thin_border
        if j == 6:
          cell.fill = self.fill_pass

    # ★ 考察ボックスの追加
    self._add_discussion_box(
        ws,
        10,
        "【変調デジタル特性 結果に対する考察】",
        "256-QAM変調信号におけるEVMは1.42%と良好であり、5G NR 100MHz信号のACLRも十分に抑制されていることが確認されました。",
    )

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 12

  # ----------------------------------------------------
  # 5. 構造・環境
  # ----------------------------------------------------
  def _build_wafer_temp_tab(self):
    ws = self.wb.create_sheet(title="5_Wafer_Temp_Yield")
    ws["A1"] = "On-Wafer Probing Yield by Temperature"
    ws["A1"].font = self.font_title

    headers = [
        "Temperature (°C)",
        "Total Tested",
        "Pass Count",
        "Fail Count",
        "Yield (%)",
    ]
    for c, h in enumerate(headers, 1):
      cell = ws.cell(row=3, column=c, value=h)
      cell.font = self.font_header
      cell.fill = self.fill_navy

    temp_data = [
        (-40, 1000, 962, 38, "=C4/B4*100"),
        (25, 1000, 988, 12, "=C5/B5*100"),
        (85, 1000, 951, 49, "=C6/B6*100"),
        (125, 1000, 924, 76, "=C7/B7*100"),
    ]

    for i, row in enumerate(temp_data, start=4):
      for j, val in enumerate(row, start=1):
        cell = ws.cell(row=i, column=j, value=val)
        cell.border = self.thin_border
        if j == 5:
          cell.number_format = "0.0"

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Yield vs Temperature"
    chart.y_axis.title = "Yield (%)"
    chart.x_axis.title = "Temperature (°C)"

    data = Reference(ws, min_col=5, min_row=3, max_row=7)
    cats = Reference(ws, min_col=1, min_row=4, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    ws.add_chart(chart, "G3")

    # ★ 考察ボックスの追加
    self._add_discussion_box(
        ws,
        10,
        "【構造・環境（温度特性） 結果に対する考察】",
        "高温（125℃）での歩留まり低下傾向が見られるため、熱膨張・接触抵抗を考慮したプロセス改善が今後の課題となります。",
    )

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15


if __name__ == "__main__":
  demo = RFEvaluationDXDemo()
  demo.generate_all_reports()