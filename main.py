import gspread

sa = gspread.service_account()
sh = sa.open("MSA 2022 Membership Form (Responses)")

wks = sh.worksheet("Form Responses 1")
