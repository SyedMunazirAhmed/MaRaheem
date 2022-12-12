import tkinter as tk
from tkinter import *
from tkinter import Menu
from tkinter import messagebox
from tkinter import ttk
import os
from tkinter import filedialog
import pandas as pd
import openpyxl

def go(event):
    global default_sheet
    cs=sheets.curselection()
    if cs:
        index1 = cs[0]
        default_sheet = event.widget.get(index1)
        display.config(text=default_sheet)
        print(default_sheet)
        return default_sheet
def callback(event):
    global default_sheet,data
    selection = event.widget.curselection()
    if selection:
        index = selection[0]
        data = event.widget.get(index)
        print(data)
        return data
def openFile():
    global df, df_list, x_axis, y_axis, canvasFlag
    global label_flag,Filepath,DatasourceName
    global prev_count, sheets,Data_source_name
    global column_name, sheet_list
    # if label_flag == True:
    #     for i in range(prev_count):
    #         column_name.destroy()
    #     label_flag = False

    filepath = filedialog.askopenfilename()
    file = open(filepath, 'r')
    # file_path["text"] = filepath
    Filepath=filepath
    file.close()
    pathname, extension = os.path.splitext(Filepath)
    DatasourceName = pathname.split('/')
    Data_source_name=DatasourceName[-1]
    print(Data_source_name)
    # df = pd.read_excel(filepath)
    # print(df.columns)
    # df_list = list(df)
    # prev_count = len(df_list)
    # df1 = pd.ExcelFile(filepath)

    df1 = pd.ExcelFile(filepath)
    sheet_names = df1.sheet_names

    for sheetlist in sheet_names:
        dfsheet = pd.read_excel(filepath, sheet_name=sheetlist)
        print(dfsheet.columns)
    DataSource.pack(side=TOP)
    sheets.pack()
    Datalbl = Label(DataSource, text=Data_source_name, height=1, width=25)
    Datalbl.pack()
    data_but.pack(side=LEFT, anchor="center", fill=X)
    data_but1.pack(side=LEFT, anchor="center", fill=X)
    # new_var=sheets.get(ANCHOR)
    # print(new_var)
    for items in sheet_names:
        sheets.insert(END, items)

    # dfs = pd.read_excel(filepath,sheet_name='People')

    sheet_list = list(sheet_names)
    # print(sheet_list)
    # f3 = dfsheet.merge(dfsheet, on="Order ID", how="left")
    # # print(f3.columns)
    # f3.to_excel("Results.xlsx", index=False)


# for items in sheet_list:
#    # #     df2 = pd.read_excel(items)
#    #     print(df2)
#      sheets = Listbox(sideframe2)
#      sheets.insert(END, items)
#    sheetsFunc(sheet_list)

def open_program():
    my_program = filedialog.askopenfilename()
    my_label.config(text=my_program)
    os.system(my_program)


def ask_qus():
    val = messagebox.askquestion("Exit?", "Do you want to exit?")
    if val == 'yes':
        root2.destroy()

def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))

    label_file["text"] = filename
    Load_excel_data()
    return None


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column)  # let the column heading = column name

    df_rows = df.to_numpy().tolist()  # turns the dataframe into a list of lists
    for row in df_rows:
        tv1.insert("", "end",
                   values=row)  # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None


def clear_data():
    tv1.delete(*tv1.get_children())
    return None


root2 = tk.Tk()
root2.title("Ocean-Book1")

root2.geometry("900x700")

my_label = Label(root2, text="")
my_label.pack(pady=20)
# create a menubar
menubar = Menu(root2)
root2.config(menu=menubar)

# create a menu
file_menu = Menu(menubar)
data_menu = Menu(menubar)
server_menu = Menu(menubar)
window_menu = Menu(menubar)
help_menu = Menu(menubar)
# add a menu item to the menu

menubar.add_cascade(
    label="File",
    menu=file_menu
)
# add the File menu to the menubar

file_menu.add_command(
    label='New',
    command=''
)
file_menu.add_command(
    label='Open',
    command=open_program
)
file_menu.add_separator()
file_menu.add_command(
    label='Paste',
    command=''
)
file_menu.add_command(
    label='Exit',
    command=ask_qus
)
menubar.add_cascade(
    label="Data",
    menu=data_menu
)
data_menu.add_command(
    label='New Data Source',
    command=lambda: File_dialog()
)
data_menu.add_separator()
data_menu.add_command(
    label='Refresh Data Source',
    command=''
)
data_menu.add_command(
    label='Duplicate Data Source',
    command=''
)
data_menu.add_command(
    label='Close Data Source',
    command=''
)
menubar.add_cascade(
    label="Server",
    menu=server_menu
)
menubar.add_cascade(
    label="Window",
    menu=window_menu
)
menubar.add_cascade(
    label="Help",
    menu=help_menu
)

f1 = tk.Frame(root2, bg="#e6e6e6", borderwidth=2, width=280, relief=SUNKEN)
f1.pack(side=LEFT, fill="y")
f1.pack_propagate(0)
data_but = Button(f1, text="Connect to Data", command=openFile)
data_but.pack(side=BOTTOM, anchor="center", fill=X)
# f2 = tk.Frame(root2, bg="#e6e6e6", borderwidth=1, height=290, relief=SUNKEN)
# f2.pack(side=BOTTOM, anchor="sw", fill="x")
# f2.pack_propagate(0)

frame1 = tk.LabelFrame(root2, text="Excel Data", height=500, width=500, )
frame1.pack(side=BOTTOM, fill="x", anchor="sw")
frame1.pack_propagate(0)
DataSource = LabelFrame(f1,height=5, width=50)
sheets = Listbox(f1, height=5, width=50)
sheets.bind("<<ListboxSelect>>",callback)

sheets.bind('<Button-3>', go)
#sheets.bind_class('<<Button-3>>', right)


# Frame for open file dialog
file_frame = tk.LabelFrame(root2, text="Open File")
file_frame.place(height=80, width=400, anchor='ne', rely=0.65, relx=0)

# Buttons
button1 = tk.Button(file_frame, text="Browse A File", )
button1.place(rely=0.65, relx=0.50)

button2 = tk.Button(file_frame, text="Load File", command=lambda: Load_excel_data())
button2.place(rely=0.65, relx=0.30)

# The file/file path text
label_file = ttk.Label(file_frame)
label_file.place()

## Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1)  # set the height and width of the widget to 100% of its container (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical",
                           command=tv1.yview)  # command means update the yaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal",
                           command=tv1.xview)  # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx.set,
              yscrollcommand=treescrolly.set)  # assign the scrollbars to the Treeview Widget
treescrollx.pack(side="bottom", fill="x")  # make the scrollbar fill the x axis of the Treeview widget
treescrolly.pack(side="right", fill="y")  # make the scrollbar fill the y axis of the Treeview widget




root2.mainloop()
