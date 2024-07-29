
import gspread
from google.oauth2.service_account import Credentials

sc  = ['https://www.googleapis.com/auth/spreadsheets']

creds = Credentials.from_service_account_file('Google Sheet API/credential.json',scopes=sc)
client = gspread.authorize(creds)

sheet_id = '1bO74tlBP21qoR3UOJNaWSA0X9iCLlb9PZdUIx14_QaI'

sheet = client.open_by_key(sheet_id)

worksheet = sheet.get_worksheet(0)

##############################

from tkinter import *
from tkinter import ttk

GUI = Tk()
GUI.geometry('500x500')
GUI.title('โปรแกรมบันทึกค่าใช้จ่ายลง Google Sheet')

FONT1 = ('Angsana New', 20)

L = ttk.Label(GUI, text='รายการค่าใช้จ่าย')
L.pack()
v_title = StringVar()
E1 = ttk.Entry(GUI, textvariable=v_title , font=FONT1, width=20)
E1.pack()
# -----------------
L = ttk.Label(GUI, text='ราคา')
L.pack()
v_price = StringVar()
E2 = ttk.Entry(GUI, textvariable=v_price , font=FONT1, width=20)
E2.pack()

def Save():
    column = worksheet.col_values(1) # Column A
    count = len(column)  + 1    # Count column A
    title = v_title.get()
    price = v_price.get()
    worksheet.update_cell(count,1,title)
    worksheet.update_cell(count,2,price)
    # Clear
    v_title.set('')
    v_price.set('')
    Update_data()

B1 = ttk.Button(GUI,text='บันทึก',command=Save)
B1.pack(ipadx=30,ipady=20,pady=20)

v_result = StringVar()
result = ttk.Label(GUI,textvariable=v_result)
result.pack()

def Update_data():
    value_list = worksheet.col_values(2)
    del value_list[0] # Clear hearder
    digit_list = [float(v) for v in value_list]
    total = sum(digit_list)
    text = 'ยอดค่าใช้จ่ายทั้งหมด {} บาท'.format(total)
    v_result.set(text)

Update_data()
GUI.mainloop()