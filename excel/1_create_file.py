from openpyxl import Workbook

wb = Workbook() # Create a new workbook
ws = wb.active # Get active sheets
ws.title = "SheetNameHere" # Rename the sheet
wb.save("test.xlsx")
wb.close()