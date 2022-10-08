from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Combobox
from tkinter import filedialog as fd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd
from screeninfo import get_monitors
import json

class App:
    
    def __init__(self):

        # Window Information
        monitor = get_monitors()
        self.window_width, self.window_height = monitor[0].width, monitor[0].height

        self.window = Tk()
        self.window.title('Test Window')
        self.window.geometry(f'{self.window_width}x{self.window_height}+0+0')
        self.canvas = Canvas(self.window, bg="blue", height=250, width=300)
        
        # Holds all Excel File Information
        self.data = None

        # Plot Information
        self.figure = plt.Figure(figsize=(6,5), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.chart_type = FigureCanvasTkAgg(self.figure, self.canvas)

        # Defaulting to Newest page of Excel
        self.option = ''
        self.options = []

    def create_button(self, x, y, text='SAMPLE BUTTON', color='blue', placeOnWindow=True):
        btn=Button(self.window, text=text, fg=color)
        if placeOnWindow: btn.place(x=x, y=y)
        return btn

    def create_label(self, x, y, text='SAMPLE LABEL', color = 'red', font = 'Helvetica', fontsize = 16):
        lbl=Label(self.window, text=text, fg=color, font=(font, fontsize))
        lbl.place(x=x, y=y)

    def create_textentry(self, x, y, text='SAMPLE TEXT ENTRY', placeOnWindow=True):
        txtfld=Entry(self.window, text=text, bd=5)
        if placeOnWindow: txtfld.place(x=x, y=y)
        return txtfld

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
        negatives['Hours Behind'] = [ahList[i] if cList[i] == 'r' else 0 for i in range(len(cList))]
        positives['Extra Hours'] = [ahList[i] if cList[i] == 'g' else 0 for i in range(len(cList))]
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
        self.chart_type.get_tk_widget().pack()

        # Filtering Data
        histogram_data['totalHours'] = histogram_data['RESERVED MTS-ESS'] + histogram_data['DC'] + histogram_data['VanCraft'] + histogram_data['CubeSmart']
        THours     = list(histogram_data['totalHours'])
        ninetyCap  = list(histogram_data['90% CAP'])
        hundredCap = list(histogram_data['100% CAP'])

        histogram_data['color'] = ['g' if THours[j] < ninetyCap[j] else 'y' if THours[j] < hundredCap[j] else 'r' for j in range(len(histogram_data))]

        # Splitting the data into 3 sets
        negatives, mehs, positives = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        cList, thList = list(histogram_data['color']), list(histogram_data['totalHours'])
        negatives['>100% CAP'] = [thList[i] if cList[i] == 'r' else 0 for i in range(len(cList))]
        mehs['90-100% CAP']    = [thList[i] if cList[i] == 'y' else 0 for i in range(len(cList))]
        positives['<90% CAP']  = [thList[i] if cList[i] == 'g' else 0 for i in range(len(cList))]
        negatives['Month/Year'], mehs['Month/Year'], positives['Month/Year'] = monthyear, monthyear, monthyear
        negatives['color'], mehs['color'], positives['color'] = ['r' for _ in histogram_data['color']], ['y' for _ in histogram_data['color']], ['g' for _ in histogram_data['color']]

        # Plotting Data
        histogram_data[['Month/Year', '90% CAP']].plot.line(legend=False, ax=self.ax, color='y').axis('off')
        histogram_data[['Month/Year', '100% CAP']].plot.line(legend=False, ax=self.ax, color='r').axis('off')

        positives.plot.bar(legend=['<90% CAP'], ax=self.ax, color=list(positives['color']), x='Month/Year').axis('on')
        mehs.plot.bar(legend=['90-100% CAP'], ax=self.ax, color=list(mehs['color']), x='Month/Year').axis('on')
        negatives.plot.bar(legend=['>100% CAP'], ax=self.ax, color=list(negatives['color']), x='Month/Year').axis('on')

        self.ax.set_title('Hours Working For Month')

    def refresh_page(self):
            
        # Creating Buttons For Visualizing Dataset
        b1 = self.create_button(10, 10, 'Hours Available For Month')
        b2 = self.create_button(10, 45, 'Hours Working For Month')
        # b3 = self.create_button(10, 80, '???')

        self.create_label(10, 115, 'Please Select A Page', 'black', fontsize=12)
        self.dd1 = self.create_dropdown(self.options, 10, 140)

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
        if '.xlsx' in self.file_directory:
            sheets = pd.ExcelFile(self.file_directory)
            self.options = sheets.sheet_names
            self.data = sheets.parse(self.options)
            self.option = sheets.sheet_names[-1]
            self.refresh_page()

    def modify_data(self, admin):
        # if not admin.registered:
        #     # messagebox.showinfo("Wait!", "You Need To Login Before You Can Use This Feature!")
        #     messagebox.showwarning(title="Wait!", message="You Need To Login Before You Can Use This Feature!")
        # else:
        #     if self.data == None:
        #         messagebox.showinfo("ShowInfo", "You Need To Login Before Committing to This Action!")
        #     else:
                # messagebox.askokcancel(title="Are You Sure?", message="This will overwrite your current excel file with modified data.\n\nContinue?")
        messagebox.showinfo(title="WIP", message="This Portion Of the Software is Currently Under Development!")

class LoginHandler(App):

    def __init__(self, window):
        self.username = ''
        self.password = ''
        self.window = window

        # True if logged in
        self.registered = False

    def logout(self):
        if not self.registered: messagebox.showinfo("Hey!", "You were never logged in! :(")
        else:
            self.registered = False
            messagebox.showinfo("Success!", "You are no longer logged in!")

    def login(self):

        if self.registered:
            messagebox.showwarning(title="Hey!", message="You're already logged in!")
            return

        self.login_canvas = Canvas(self.window, bg="blue", height=250, width=300)
        self.login_canvas.pack()

        # Username and Password text entry w/ Login Button
        self.user_entry = Entry(self.login_canvas, text='Username Here', bd=5)
        self.pass_entry = Entry(self.login_canvas, show='*', text='Password Here', bd=5)
        submit = Button(self.login_canvas, text='Login', fg='blue')
        cancel = Button(self.login_canvas, text='Cancel', fg='blue')

        self.user_entry.place(x=25, y=25)
        self.pass_entry.place(x=25, y=60)
        submit.place(x=25, y=100)
        cancel.place(x=100, y=100)

        # Gather user and pass then encrypt the data, compare to data
        submit.bind('<Button-1>', self.verify_login)
        cancel.bind('<Button-1>', self.close_window)

    def close_window(self, _): self.login_canvas.destroy()

    def verify_login(self, _):

        def encrypt(user, passwd):

            logins = pd.read_json('database.json')

            encrypted_user = ''
            encrypted_pass = ''
            for u in user: encrypted_user += chr((ord(u) + 15) * 5 - 33)
            for p in passwd: encrypted_pass += chr((ord(p) + 18) * 4 - 17)

            for login in range(len(logins)):
                if logins['username'][login] == encrypted_user and logins['password'][login] == encrypted_pass:
                    self.user_entry.delete(0, len(encrypted_user))
                    self.pass_entry.delete(0, len(encrypted_pass))
                    self.login_canvas.destroy()
                    messagebox.showinfo("Success!", "You're all set up!")
                    self.registered = True
            
            if not self.registered:
                messagebox.showwarning(title="Oops!", message="Your Username or Password seem to be incorrect!")
                self.pass_entry.delete(0, len(encrypted_pass))


        encrypt(self.user_entry.get(), self.pass_entry.get())

    def create_account(self):
        '''
            Ask for admin password
                Create Username and Password
                Verify Password
                Email

            Forgot Password?
                Sends email
                Creates Dummy Login with same username but password sent to email
                Once logged in with dummy password, remove from database and replace old password with new one.
        '''
        pass

    def forgot_password(self):
        # Promps with username and sends email
        # Has Forgot Username
        # If Forgot Username, Inputs email and sends an email
        pass

def first_time_run(app, admin):
    
    def create_admin_account(_):
        # Encrypt Data

        encrypted_email, encrypted_password, encrypted_confirmed_password = '', '', '' 
        for i in ADMIN_EMAIL.get(): encrypted_email += chr((ord(i) + 15) * 5 - 33)
        for i in ADMIN_PASSWORD.get(): encrypted_password += chr((ord(i) + 18) * 4 - 17)
        for i in ADMIN_CONFIRM_PASSWORD.get(): encrypted_confirmed_password += chr((ord(i) + 18) * 4 - 17)

        if encrypted_email == '' or encrypted_password == '' or encrypted_confirmed_password == '':
            messagebox.showwarning(title="Hey!", message="Please Fill Out Every Form!")
        else:
            if encrypted_password != encrypted_confirmed_password:
                ADMIN_PASSWORD.delete(0, len(encrypted_password))
                ADMIN_CONFIRM_PASSWORD.delete(0, len(encrypted_confirmed_password))
                messagebox.showwarning(title="Oops!", message="Your Passwords don't match!")

            else:
                ADMIN_EMAIL.delete(0, len(encrypted_email))
                ADMIN_PASSWORD.delete(0, len(encrypted_password))
                ADMIN_CONFIRM_PASSWORD.delete(0, len(encrypted_confirmed_password))
                messagebox.showinfo("Account Successfully Created!", "Enjoy The Program!")
                app.canvas.destroy()

                admin_account = json.dumps([{'username': encrypted_email, 'password': encrypted_password, 'email': encrypted_email}])
                with open('./database.json', 'w') as file: file.write(admin_account)

                # First Time run opens menu
                menu1 = app.create_window_dropdown('Menu')
                app.create_tab_for_menu(menu1, 'Open...', app.open_file)
                app.create_tab_for_menu(menu1, 'Login',   admin.login)
                app.create_tab_for_menu(menu1, 'Logout',  admin.logout)
                # app.create_tab_for_menu(menu1, 'Modify',  lambda event=admin: app.modify_data(event))

    # Making Admin Account
    a_email = Label(app.canvas, text="E-Mail\n/Username:", fg='black', font=("Helvetica", 12))
    a_passwd = Label(app.canvas, text="Password:", fg='black', font=("Helvetica", 12))
    a_cpasswd = Label(app.canvas, text="Confirm\nPassword:", fg='black', font=("Helvetica", 12))
    
    ADMIN_EMAIL = Entry(app.canvas, text='E-Mail', bd=5, fg='purple', justify='center')
    ADMIN_PASSWORD = Entry(app.canvas, show='*', text='Password', bd=5, justify='center')
    ADMIN_CONFIRM_PASSWORD = Entry(app.canvas, show='*', text='ConfirmPassword', bd=5, justify='center')

    a_email.place(x=15, y=15)
    ADMIN_EMAIL.place(x=110, y=15)
    a_passwd.place(x=15, y=70)
    ADMIN_PASSWORD.place(x=110, y=70)
    a_cpasswd.place(x=15, y=105)
    ADMIN_CONFIRM_PASSWORD.place(x=110, y=105)

    submit = Button(app.canvas, text='Create Admin Login', fg='blue')
    submit.place(x=75, y=150)

    submit.bind('<Button-1>', create_admin_account)

    app.canvas.pack()

def main():
    app = App()
    admin = LoginHandler(app.window)

    first_run = True
    with open('database.json', 'r') as file:
        for i in file:
            if i.strip() != "[]":
                first_run = False
                break

    if first_run: first_time_run(app, admin)
    else:
        menu1 = app.create_window_dropdown('Menu')
        app.create_tab_for_menu(menu1, 'Open...', app.open_file)
        app.create_tab_for_menu(menu1, 'Login',   admin.login)
        app.create_tab_for_menu(menu1, 'Logout',  admin.logout)
        # app.create_tab_for_menu(menu1, 'Modify',  lambda event=admin: app.modify_data(event))

    app.window.mainloop()

if __name__ == '__main__':
    main()

# ***IMPORTANT*** | RENAME MAIN.PY TO THE APPLICATION NAME.
# pyinstaller --onefile main.py