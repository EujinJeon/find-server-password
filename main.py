import sys
import os

import xlwings as xw
import pandas as pd

import tkinter as tk
import tkinter.filedialog
import tkinter.font

import clipboard as cp

window = tk.Tk()
window.title("Server Password Tool")
window.geometry("440x300+100+100")
window.resizable(False, False)
myfont1 = tk.font.Font(family="맑은 고딕", size=9)

filename = ""
server_passwd = ""

def openfile():
    global filename
    filename = tk.filedialog.askopenfilename()


def open_excel(passwd, server_ip):
  global server_passwd
  #print(filename)
  try :
    app = xw.App(visible=False)
    wb = xw.Book(filename, password=passwd)
    #print(passwd)
    #print(server_ip)
    sheet1 = wb.sheets['sheet1'] #Set Server List sheet
    sheet2 = wb.sheets['sheet2'] #Set Passwd List sheet
  
    #Server List Sheet
    s_last_row = sheet1.range('A' + str(sheet1.cells.last_cell.row)).end('up').row
    s_column_values = sheet1.range(f'A1:d{s_last_row}').value
    
    #input_str = input()
    input_str = server_ip
    s_matching_rows = [cell[1] for cell in enumerate(s_column_values) if input_str in str(cell)]
    
    if len(s_matching_rows) == 0:
        server_passwd = "존재하지 않는 서버입니다."
        wb.close()
        app.kill()
        return
      
    service_name = s_matching_rows[0][0]
    # Passwd Sheet
    p_last_row = sheet2.range('A' + str(sheet2.cells.last_cell.row)).end('up').row
    p_column_values = sheet2.range(f'A1:B{p_last_row}').value
    p_matching_rows = [cell[1] for cell in enumerate(p_column_values) if service_name in str(cell)]
    wb.close()
    app.kill()
    server_passwd = p_matching_rows[0][1]
    
    #print(server_passwd)
    cp.copy(server_passwd)
    
  except :
        return

# 1 행
label10 = tk.Label(
    window,
    text="데이터 베이스 경로",
    font=myfont1,
    bg="white",
    fg="black",
    height=1,
    width=15,
)
label10.grid(row=0, column=0, padx=5, pady=10)
open_button = tk.Button(
    window,
    text="파일 열기",
    overrelief="solid",
    width=10,
    command=openfile,
    repeatdelay=1000,
    repeatinterval=100,
)
open_button.grid(row=0, column=2, padx=5, pady=10)

# 2 행
label20 = tk.Label(
    window, text="DB패스워드", font=myfont1, bg="white", fg="black", height=1, width=15
)
label20.grid(row=1, column=0, padx=5, pady=10)
passwd_ent = tk.Entry(window, width=30, show="*")
passwd_ent.grid(row=1, column=1, padx=5, pady=10)


# 3 행

label30 = tk.Label(
    window, text="호스트IP", font=myfont1, bg="white", fg="black", height=1, width=15
)
label30.grid(row=2, column=0, padx=5, pady=10)
server_ent = tk.Entry(window, width=30)
server_ent.grid(row=2, column=1, padx=5, pady=10)

button23 = tk.Button(
    window,
    text="검색",
    overrelief="solid",
    width=10,
    command=lambda :open_excel(passwd_ent.get(), server_ent.get()),
    repeatdelay=1000,
    repeatinterval=100,
)
button23.grid(row=3, column=1, padx=5, pady=10)

# 4 행
label40 = tk.Label(
    window, text="호스트 비밀번호", font=myfont1, bg="white", fg="black", height=1, width=15
)
label40.grid(row=4, column=0, padx=5, pady=10)


while True:
    f_label = tk.Label(
        window, text=filename, font=myfont1, bg="white", fg="black", height=1, width=30
    )
    f_label.grid(row=0, column=1, padx=5, pady=10)
    
    label41 = tk.Label(
    window, text=server_passwd, font=myfont1, bg="white", fg="black", height=1, width=30
    )
    label41.grid(row=4, column=1, padx=5, pady=10)
    window.update()

    try:
        f_label.destroy()
        label41.destroy()
    except:
        break

window.mainloop()
