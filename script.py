import gspread

sa = gspread.service_account()
sh = sa.open("Image to Sheet")

wks = sh.worksheet("Sheet1")

wks.update("A3", "hey0o")
# print(wks.acell("A1").value)
# or do print(wks.cell(1,1).value)

