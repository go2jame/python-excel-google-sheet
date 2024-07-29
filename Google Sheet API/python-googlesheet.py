# pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib gspread

import gspread
from google.oauth2.service_account import Credentials

sc  = ['https://www.googleapis.com/auth/spreadsheets']

creds = Credentials.from_service_account_file('Google Sheet API/credential.json',scopes=sc)
client = gspread.authorize(creds)

sheet_id = '1bO74tlBP21qoR3UOJNaWSA0X9iCLlb9PZdUIx14_QaI'

sheet = client.open_by_key(sheet_id)

worksheet = sheet.get_worksheet(0)
result = worksheet.acell('B5').value
result2 = worksheet.cell(1,1).value

cell = worksheet.acell('B1', value_render_option='FORMULA').value

value_list = worksheet.row_values(1)
print(value_list)

column = worksheet.col_values(1)
print(len(column))


# print(cell)
# print(result)
# print(result2)

# worksheet.update_cell(1,1,'Update from my computer')
# worksheet.update_acell('B1','=A1&"สวัสดี ส่งข้อมูลจากคอม"')