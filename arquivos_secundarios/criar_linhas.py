from openpyxl import load_workbook

wb = load_workbook("arquivo_base_qsgr.xlsx")

ws_qsgr = wb["QSGR"]

linha = 74

for i in range(80):
    ws_qsgr.insert_rows(idx=linha + i)
wb.save("arquivo_base_novas_linhas.xlsx")