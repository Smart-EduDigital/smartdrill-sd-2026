import openpyxl
wb = openpyxl.load_workbook("C:/Users/MIP-USER/Documents/Project Soal/Quiz_Web_Matematika_Paket_1.xlsx")
ws = wb.active
print("Total rows now in Quiz_Web_Matematika_Paket_1.xlsx:", ws.max_row - 1)