from tkinter import *

app = Tk()

filename_text = StringVar()
filename_label = Label(app, text='File name:', font=('bold', 14), pady=20)
filename_label.grid(row=0, column=0, sticky=W)
filename_entry = Entry(app, textvariable=filename_text)
filename_entry.grid(row=0, column=1)

app.title("Spreadsheet Manager")
app.geometry('700x700')

app.mainloop()
