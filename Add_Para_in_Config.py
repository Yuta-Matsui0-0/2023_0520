import os
import openpyxl

# Inputフォルダのパス
input_folder_path = "./Input"
# Outputフォルダのパス
output_folder_path = "./Output"

# Inputフォルダ内の全Excelファイルを処理
for file_name in os.listdir(input_folder_path):
    if not file_name.endswith(".xlsx"):
        continue

    # Excelファイルを読み込む
    input_file_path = os.path.join(input_folder_path, file_name)
    wb = openpyxl.load_workbook(input_file_path)

    # "ABC"シートのB1セルを確認し、条件に合致した場合は文言を追加する
    sheet_name = "ABC"
    if sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        if sheet["B1"].value == "TestSample":
            sheet["B1"].value += "n=0-1"

    # 出力ファイル名を決定し、Excelファイルを保存する
    output_file_path = os.path.join(output_folder_path, file_name.replace(".xlsx", "_AddPara.xlsx"))
    wb.save(output_file_path)
