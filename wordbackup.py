import random
from tkinter import *
import tkinter
import xlwings as xw


global loc,canvas_obj1,canvas_obj2
loc = 1

def on_resize(evt):
    tk.configure(width=evt.width, height=evt.height)
    canvas.create_rectangle(0, 0, canvas.winfo_width(), canvas.winfo_height(), fill=TRANSCOLOUR, outline=TRANSCOLOUR)

def refreshText():
    global  loc,canvas_obj1,canvas_obj2
    app = xw.App(visible=False, add_book=False)  # xlwings不打开文件的情况下操作
    book = app.books.open("word8000.xls")  # file_path 为文件路径
    sheet = book.sheets['Sheet1']  # sheet_name 为表名称

    loc = random.randint(1, sheet.range(1, 1).end('down').row + 1)
    text1 = sheet.range(loc, 1).value
    text2 = sheet.range(loc, 2).value
    # print(text1, text2)
    book.close()
    app.quit()

    canvas.delete(canvas_obj1)
    canvas.delete(canvas_obj2)

    canvas_obj1 = canvas.create_text(200, 30, text=text1, font=('Times New Roman', 30), fill='DarkOrange',
                                     justify=CENTER)
    canvas_obj2 = canvas.create_text(200, 70, text=text2, font=('Times New Roman', 15), fill='PaleVioletRed',
                                     justify=CENTER)

    tk.update()

    tk.after(5000, refreshText)

if __name__ == '__main__':
    TRANSCOLOUR = 'gray'
    tk = Tk()
    tk.geometry('400x100+1000+100')
    tk.title('')
    tk.iconphoto(False, tkinter.PhotoImage(file='R-C.png'))
    tk.wm_attributes('-transparentcolor', TRANSCOLOUR)
    tk.wm_attributes("-topmost", 1)
    tk.overrideredirect(True)
    tk.bind('<Configure>', on_resize)

    canvas = Canvas(tk)
    tk.winfo_width()
    canvas.pack(fill=BOTH, expand=Y)
    # canvas.config(highlightthickness=0)
    tk.update()

    app = xw.App(visible=False, add_book=False)  # xlwings不打开文件的情况下操作
    book = app.books.open("word8000.xls")  # file_path 为文件路径
    sheet = book.sheets['Sheet1']  # sheet_name 为表名称

    loc = random.randint(1, sheet.range(1, 1).end('down').row + 1)
    text1 = sheet.range(loc, 1).value
    text2 = sheet.range(loc, 2).value
    # print(text1, text2)

    book.close()
    app.quit()

    canvas_obj1 = canvas.create_text(200, 30, text=text1, font=('Times New Roman', 30), fill='DarkOrange',
                                     justify=CENTER)
    canvas_obj2 = canvas.create_text(200, 70, text=text2, font=('Times New Roman', 15), fill='PaleVioletRed',
                                     justify=CENTER)

    tk.after(5000,refreshText)

    tk.mainloop()