import os
import pandas as pd
import datetime
from openpyxl import Workbook

# 検索ディレクトリと出力ディレクトリのパス
input_dir = 'path_to_your_input_directory'
output_dir = 'path_to_your_output_directory'

# 新しいワークブックを作成
wb = Workbook()

# ヘッダを作成
ws = wb.active
ws.append(["File name", "A_sample", "", "B_sample", ""])
ws.append(["", "シートの有無", "情報の有無", "シートの有無", "情報の有無"])

# 入力フォルダ内の全てのExcelファイルを走査
for file_name in os.listdir(input_dir):
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        file_path = os.path.join(input_dir, file_name)

        try:
            # A_sampleシートの確認
            a_sheet = pd.read_excel(file_path, sheet_name="A_sample")
            a_sheet_exists = "v"
            a_data_exists = "v" if pd.notna(a_sheet.loc[2, "C"]) else "-"
        except Exception:
            a_sheet_exists = "-"
            a_data_exists = "-"

        try:
            # B_sampleシートの確認
            b_sheet = pd.read_excel(file_path, sheet_name="B_sample")
            b_sheet_exists = "v"
            b_data_exists = "v" if pd.notna(b_sheet.loc[2, "C"]) else "-"
        except Exception:
            b_sheet_exists = "-"
            b_data_exists = "-"

        # 結果をワークシートに追加
        ws.append([file_name, a_sheet_exists, a_data_exists, b_sheet_exists, b_data_exists])

# ファイル名を現在の日付・時間に基づいて生成
output_file_name = "Check_Sheet_" + datetime.datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"

# ファイルを出力ディレクトリに保存
output_file_path = os.path.join(output_dir, output_file_name)
wb.save(output_file_path)
