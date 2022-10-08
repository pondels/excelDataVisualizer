from tkinter import *
from tkinter.ttk import Combobox
from tkinter import filedialog as fd
import pandas as pd

class App:
    
    def __init__(self):
        self.window = Tk() 
        self.window.title('Test Window')
        self.window.geometry('1280x720+0+0')
        
        self.data = None

    def create_button(self, x, y, text='SAMPLE BUTTON', color='blue'):
        btn=Button(self.window, text=text, fg=color)
        btn.place(x=x, y=y)
        btn.bind('<Button-1>', self.buttonClicked())

    def create_label(self, x, y, text='SAMPLE LABEL', color = 'red', font = 'Helvetica', fontsize = 16):
        lbl=Label(self.window, text=text, fg=color, font=(font, fontsize))
        lbl.place(x=x, y=y)

    def create_textentry(self, x, y, text='SAMPLE TEXT ENTRY'):
        txtfld=Entry(self.window, text=text, bd=5)
        txtfld.place(x=x, y=y)

    def create_dropdown(self, options, x, y):
        var = StringVar()
        var.set(options[1])
        cb=Combobox(self.window, values=options)
        cb.place(x=x, y=y)

    def create_choice(self, options, x, y, mode = 'multiple'):
        lb=Listbox(self.window, height=5, selectmode=mode)
        for option in options: lb.insert(END, option)
        lb.place(x=x, y=y)

    def create_radiobutton(self, x, y, text='SAMPLE RADIOBUTTON'):
        v0=IntVar()
        r1=Radiobutton(self.window, text=text, variable=v0,value=1)
        r1.place(x=x,y=y)
                        
    def create_checkbutton(self, x, y, text='SAMPLE CHECKBUTTON'):
        v1 = IntVar()
        C1 = Checkbutton(self.window, text = text, variable = v1)
        C1.place(x=x, y=y)
        
    def create_window_dropdown(self, label):
        m = Menu(self.window)
        self.window.config(menu=m)
        file_menu = Menu(m, tearoff=False)
        m.add_cascade(label=label, menu=file_menu)
        return file_menu

    def create_tab_for_menu(self, file_menu, label, function):
        file_menu.add_command(label=label, command=function)

    def visualize_data(self):
        
        main = self.data[7:]
        header = ['Month/Year', '100% CAP', '90% CAP', 'RESERVED MTS-ESS', 'DC', 'VanCraft', 'CubeSmart', 'Available Hours', 'Confirmed Hours', 'Confirmed Details', 11, 12, 13, 14, 15, 16]
        main.columns = header
        
        main = main.drop([11, 12, 13, 14, 15, 16, 'Confirmed Details'], axis=1)
        main = main.drop(7, axis=0)
        # print(main)
        print(main.head())

    # *** TRIGGERED FUNCTIONS *** #
    def open_file(self):
        self.file_directory = fd.askopenfilename(title="Open a File", filetypes=(("xlxs files", ".*xlsx"), ("csv files", "*.*csv")))
        if '.xlsx' in self.file_directory:
            self.data = pd.read_excel(
                io=self.file_directory,
                engine='openpyxl'
            )
            self.visualize_data()



def main():
    app = App()

    # app.create_button(10, 10, 'open file')
    menu1 = app.create_window_dropdown('Menu')
    app.create_tab_for_menu(menu1, 'Open...', app.open_file)
    # app.create_label(60, 50)
    # app.create_textentry(80, 150)
    # app.create_dropdown(("one", "two", "three", "four"), 60, 150)
    # app.create_choice(("one", "two", "three", "four"), 250, 150)
    # app.create_radiobutton(100, 50, 'male')
    # app.create_radiobutton(180, 50, 'female')
    # app.create_checkbutton(100, 100, 'Cricket')
    # app.create_checkbutton(180, 100, 'Tennis')
    # app.create_button(50, 50, 'TEST')
    app.window.mainloop()

if __name__ == '__main__':
    main()

# ***IMPORTANT*** | RENAME MAIN.PY TO THE APPLICATION NAME.
# pyinstaller --onefile main.py