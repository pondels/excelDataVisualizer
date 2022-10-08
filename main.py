from tkinter import *
from tkinter.ttk import Combobox
from tkinter import filedialog as fd
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd

class App:
    
    def __init__(self):
        self.window = Tk() 
        self.window.title('Test Window')
        self.window.geometry('1920x1080+0+0')

        self.canvas = Canvas(self.window, bg="blue", height=250, width=300)
        
        # Holds all Excel File Information
        self.data = None

        # Plot Information
        self.figure = plt.Figure(figsize=(6,5), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.chart_type = FigureCanvasTkAgg(self.figure, self.canvas)

        # Defaulting to Newest page of Excel
        self.option = ''

    def create_button(self, x, y, text='SAMPLE BUTTON', color='blue'):
        btn=Button(self.window, text=text, fg=color)
        btn.place(x=x, y=y)
        return btn

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
        return cb

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

    def gather_hours_data(self):
        main_df = self.data[self.option][7:]
        header = ['Month/Year', '100% CAP', '90% CAP', 'RESERVED MTS-ESS', 'DC', 'VanCraft', 'CubeSmart', 'Available Hours', 'Confirmed Hours', 'Confirmed Details', 11, 12, 13, 14, 15, 16]
        main_df.columns = header
        
        main_df = main_df.drop([11, 12, 13, 14, 15, 16, 'Confirmed Details'], axis=1)
        main_df = main_df.drop(7, axis=0)
        return main_df

    # *** TRIGGERED FUNCTIONS ***

    def AvailableHoursChart(self, _):

        # Display Information Based On Sheet Selected
        sheet = self.dd1.get()
        if sheet != '': self.option = sheet

        # Gathering Data
        hours_data = self.gather_hours_data()
        histogram_data = hours_data[['Month/Year', 'Available Hours']].reset_index()
        
        self.refresh_page()

        # Creating Graph Template
        self.chart_type.get_tk_widget().pack()
        
        # Filtering Data
        histogram_data['color'] = ['r' if list(histogram_data['Available Hours'])[j] < 0 else 'g' for j in range(len(histogram_data))]
        ahList, cList = list(histogram_data['Available Hours']), list(histogram_data['color'])

        # Splitting Data
        negatives, positives = pd.DataFrame(), pd.DataFrame()
        negatives['Available Hours'] = [ahList[i] if cList[i] == 'r' else 0 for i in range(len(cList))]
        positives['Available Hours'] = [ahList[i] if cList[i] == 'g' else 0 for i in range(len(cList))]
        negatives['color'], positives['color'] = ['r' for _ in histogram_data['color']], ['g' for _ in histogram_data['color']]
        negatives['Month/Year'], positives['Month/Year'] = histogram_data['Month/Year'], histogram_data['Month/Year']

        # Plotting Data
        negatives.plot.bar(legend=True, ax=self.ax, color=list(negatives['color']), x='Month/Year')
        positives.plot.bar(legend=True, ax=self.ax, color=list(positives['color']), x='Month/Year')
        self.ax.set_title('Total Hours Needed For Month')

    def capacityChart(self, _):

        # Display Information Based On Sheet Selected
        sheet = self.dd1.get()
        if sheet != '': self.option = sheet

        # Gathering Data
        hours_data = self.gather_hours_data()
        histogram_data = hours_data[['Month/Year', '100% CAP', '90% CAP', 'RESERVED MTS-ESS', 'DC', 'VanCraft', 'CubeSmart']].reset_index()

        monthyear = histogram_data['Month/Year']

        self.refresh_page()

        # Creating Graph Template
        # figure = plt.Figure(figsize=(6,5), dpi=100)
        # ax = self.figure.add_subplot(111)
        # chart_type = FigureCanvasTkAgg(figure, self.window)
        # chart_type.get_tk_widget().pack()
        self.chart_type.get_tk_widget().pack()

        # Filtering Data
        histogram_data['totalHours'] = histogram_data['RESERVED MTS-ESS'] + histogram_data['DC'] + histogram_data['VanCraft'] + histogram_data['CubeSmart']
        THours = list(histogram_data['totalHours'])
        ninetyCap = list(histogram_data['90% CAP'])
        hundredCap = list(histogram_data['100% CAP'])

        histogram_data['color'] = ['g' if THours[j] < ninetyCap[j] else 'y' if THours[j] < hundredCap[j] else 'r' for j in range(len(histogram_data))]

        # Splitting the data into 3 sets
        negatives, mehs, positives = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        negatives['totalHours'] = [histogram_data['totalHours'][i] if histogram_data['color'][i] == 'r' else 0 for i in range(len(histogram_data['color']))]
        mehs['totalHours']      = [histogram_data['totalHours'][i] if histogram_data['color'][i] == 'y' else 0 for i in range(len(histogram_data['color']))]
        positives['totalHours'] = [histogram_data['totalHours'][i] if histogram_data['color'][i] == 'g' else 0 for i in range(len(histogram_data['color']))]
        negatives['Month/Year'], mehs['Month/Year'], positives['Month/Year'] = monthyear, monthyear, monthyear
        negatives['color'], mehs['color'], positives['color'] = ['r' for _ in histogram_data['color']], ['y' for _ in histogram_data['color']], ['g' for _ in histogram_data['color']]

        # Plotting Data
        positives.plot.bar(legend=True, ax=self.ax, color=list(positives['color']), x='Month/Year')
        mehs.plot.bar(legend=True, ax=self.ax, color=list(mehs['color']), x='Month/Year')
        negatives.plot.bar(legend=True, ax=self.ax, color=list(negatives['color']), x='Month/Year')
        
        histogram_data[['Month/Year', '100% CAP']].plot.line(legend=True, ax=self.ax, color='r', x='Month/Year')
        histogram_data[['Month/Year', '90% CAP']].plot.line(legend=True, ax=self.ax, color='y', x='Month/Year')
        self.ax.set_title('Hours Working For Month')

    def refresh_page(self):
        sheets = pd.ExcelFile(self.file_directory)
            
        self.data = sheets.parse(sheets.sheet_names)
        self.option = sheets.sheet_names[-1]
            
        # Creating Buttons For Visualizing Dataset
        b1 = self.create_button(10, 10, 'Hours Available For Month')
        b2 = self.create_button(10, 45, 'Hours Working For Month')
        b3 = self.create_button(10, 80, '???')

        self.create_label(10, 115, 'Please Select A Page', 'black', fontsize=12)
        self.dd1 = self.create_dropdown(sheets.sheet_names, 10, 140)

        # Creating Functions for the buttons
        b1.bind('<Button-1>', self.AvailableHoursChart)
        b2.bind('<Button-1>', self.capacityChart)
        # b3.bind('<Button-1>', self.visualize_data)

        # Resetting Plotting Window
        self.canvas.destroy()
        self.canvas = Canvas(self.window, bg="blue", height=250, width=300)
        self.canvas.pack()

        self.figure = plt.Figure(figsize=(6,5), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.chart_type = FigureCanvasTkAgg(self.figure, self.canvas)

    def open_file(self):
        self.file_directory = fd.askopenfilename(title="Open a File", filetypes=(("xlxs files", ".*xlsx"), ("csv files", "*.*csv")))
        if '.xlsx' in self.file_directory: self.refresh_page()

def main():
    app = App()
    menu1 = app.create_window_dropdown('Menu')
    app.create_tab_for_menu(menu1, 'Open...', app.open_file)
    app.window.mainloop()

if __name__ == '__main__':
    main()

# ***IMPORTANT*** | RENAME MAIN.PY TO THE APPLICATION NAME.
# pyinstaller --onefile main.py

# TODO
'''
    Grab Window Size and use that for dimensions
    Get graphs to reset
    Fix graph displays
        Axis
    (Optional) One more visual (if necessary)

    *** Stretch Goals ***

    Import CSV files
    Use the CSV files to update the newest excel page

'''