import os
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter
from docx import Document

# InputとOutputフォルダのパス
input_directory = 'path_to_your_input_directory'
output_directory = 'path_to_your_output_directory'

# ファイルの読み取り
files = os.listdir(input_directory)
word_files = [f for f in files if f.endswith('.docx')]

# Excelファイル名の準備
now = datetime.now()
output_filename = f'AnchorChecker_{now.strftime("%Y%m%d-%H%M%S")}.xlsx'
output_filepath = os.path.join(output_directory, output_filename)

# Excelファイルの準備
writer = pd.ExcelWriter(output_filepath, engine='openpyxl')

summary = []

for file in word_files:
    # 図の検出とナンバリング、アンカーの検出
    document = Document(os.path.join(input_directory, file))
    figures = []
    anchor_figures = 0
    for i, rel in enumerate(document.part.rels.values()):
        if "image" in rel.reltype:
            figures.append(i+1)
            if rel.reltype == "anchor":
                anchor_figures += 1
    data = pd.DataFrame(figures, columns=["図のナンバー"])
    data["アンカーの有無"] = data["図のナンバー"].apply(lambda x: "有" if x <= anchor_figures else "無")

    # 情報のExcelファイルへの出力
    data.to_excel(writer, sheet_name=file, index=False)

    # Summary用のデータ作成
    summary.append([file, len(figures), anchor_figures])

# Summaryシートの作成
summary_df = pd.DataFrame(summary, columns=["ファイル名", "図の個数", "アンカー付きの図の個数"])
summary_df.to_excel(writer, sheet_name="Summary", index=False)

# シート間リンクの作成
workbook = writer.book
summary_sheet = workbook["Summary"]
for row in range(2, len(summary_df) + 2):
    file_cell = summary_sheet.cell(row=row, column=1)
    file_cell.value = f'=HYPERLINK("#{file_cell.value}!A1", "{file_cell.value}")'

writer.save()
