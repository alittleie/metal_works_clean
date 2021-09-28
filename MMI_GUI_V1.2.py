import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
import pandas as pd
from tkinter import messagebox
import xlsxwriter
import warnings


LARGE_FONT= ("Verdana", 12)
class holdy_boi(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)

        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, PageOne, PageTwo):
            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartPage)

    def show_frame(self, cont):
        frame = self.frames[cont]

       # frame.config(highlightbackground="black", highlightthickness=2)

        frame.tkraise()

class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
        # label = tk.Label(self, text="Input Data", font=LARGE_FONT)
        # label.pack(pady=10,padx=10)

        image = tk.PhotoImage(file="1420310-logo-720w_25.gif")


        label = Label(image=image)
        label.image = image  # keep a reference!
        label.pack()
        label.place(x=75, y=235)



        S = tk.Scrollbar()
        global T
        T = tk.Text( height=1, width=40,font =("Verdana", 8))
        T.pack()
        T.place(x=8,y=95)

        global T2
        T2 = tk.Text(height=1, width=40, font=("Verdana", 8))
        T2.pack()
        T2.place(x=8, y=165)
        global Tchoice
        Tchoice = StringVar(self)
        Tchoice.set('Choose a File Type for Reformatting')

        # Dictionary with options
        choices = {'                 Scale Ticket Data                   ',
                   '                 Example 1                           ' ,
                   '                 Example 2                           '}
        #Tchoice.set('Pizza')  # set the default option
        #8db3be
        popupMenu = OptionMenu(self, Tchoice, *choices)
        #flat, groove, raised, ridge, solid, or sunken
        popupMenu.config(bg='#d1d1e8',activebackground = '#d1d1e8',   )
        popupMenu.pack()
        popupMenu.place(x=30, y = 10)



        button3 = tk.Button(self, text="Load File",
                            command=lambda: self.load_file(), bg = '#d1d1e8')
        button3.pack()
        button3.place(x=120, y=60)
        button4 = tk.Button(self, text="Choose Save Location",
                            command=lambda: self.choose_location(), bg = '#d1d1e8')
        button4.pack()
        button4.place(x=85, y=125)


        button5 = tk.Button(self, text="Reformat",
                            command=lambda: self.reformat(),bg='#7691cb', fg='black')
        button5.pack()
        button5.place(x=118, y=195)


    def load_file(self):
        global string_name
        global fname
        T.delete('1.0', END)
        fname = tk.filedialog.askopenfile(mode='rb', title='Choose a file')
        string_name= str(fname)
        string_name = string_name.split('name=')
        string_name = string_name[1].strip('>')
        string_name = string_name.strip('\'')
        T.insert(tk.END, string_name  )



    def choose_location(self):
        loc = filedialog.asksaveasfilename(initialdir = "/",title = "Select Save Location",filetypes = (("xlsx files","*.xlsx"),))
        global string_name2
        name = loc
        T2.delete('1.0', END)
        string_name2 = str(name)
        if '.xlsx' not in string_name2:
            string_name2 = string_name2+ '.xlsx'
        T2.insert(tk.END, string_name2)



    def reformat(self):
        if Tchoice.get() == 'Choose a File Type for Reformatting':
            tk.messagebox.showerror(title='Error', message='Please Select Data Type')
        bool1 = False
        bool2 = False
        bool3 = True
        try:
            print(" ")
            print(T.get("1.0", "end-1c"))
            bool1 = True

        except NameError:
            tk.messagebox.showerror(title='Error', message= 'Input File Needed')
        try:
            print(T2.get("1.0", "end-1c"))
            bool2 = True
        except NameError:
            tk.messagebox.showerror(title='Error', message= 'Output Destination Needed')
        try:
            if T.get("1.0", "end-1c") == T2.get("1.0", "end-1c"):
                bool3 = False
                tk.messagebox.showerror(title='Error', message='Input and Output Destination is the Same' )
        except:
            pass

        if Tchoice.get() == '                 Scale Ticket Data                   ':
            if bool1 == True and bool2 == True and bool3 == True:

                try:
                    df = pd.read_excel( T.get("1.0", "end-1c") ,
                                    sheet_name='Sheet1')
                    print(" ")
                    print('Reformatting Data...\nDepending on size may take several minutes...')
                except:
                    tk.messagebox.showerror(title='Error', message='Data Input Error')
                try:
                    df2 = pd.read_excel( T.get("1.0", "end-1c") ,
                                    sheet_name='Sheet2')
                    dfbool = True
                except:
                    dfbool = False

                if dfbool== True:
                    df = pd.concat([df,df2])

                df.dropna(axis=1, how='all', inplace=True)


                if len(df.columns) not in [21,19,22,20]:
                    tk.messagebox.showerror(title='Error', message='Non-Standard Format Error-Contact Alex')
                
                if len(df. columns) == 21:
                    df = df.drop(df.columns[[1, 2, 4, 6, 8, 9, 10, 12, 13, 14, 15, 17, 18, 19]], axis=1)
                elif len(df. columns) == 19:
                    df = df.drop(df.columns[[1, 2, 4, 6, 8, 9, 11, 12, 13, 15,16, 17]], axis=1)
                elif len(df. columns) == 20:
                    df = df.drop(df.columns[[1, 2, 4, 6, 8, 9, 10,11, 12, 14,16, 17,18]], axis=1)
                elif len(df. columns) == 22:
                    df = df.drop(df.columns[[1, 2, 4, 6,8, 9,10, 12,13,14, 15, 17,18,19,21]], axis=1)
                df = df.iloc[9:, :]

                df.columns = ['Void', 'Ticket ID', 'Ticket Date', 'Company', 'Container #', 'Area', 'Weight']
                df['Area'] = df['Area'].shift(axis=0, periods=-1)
                df['Weight'] = df['Weight'].shift(axis=0, periods=-1)
                df.dropna(axis=0, how='all', inplace=True)

                df['Ticket ID'] = df['Ticket ID'].fillna(method='ffill', axis=0)
                df['Ticket Date'] = df['Ticket Date'].fillna(method='ffill', axis=0)
                df['Company'] = df['Company'].fillna(method='ffill', axis=0)

                void_new = []
                id_hold = -99999
                for i in range(len(df)):
                    if df['Ticket ID'].iloc[i] == id_hold:
                        void_new.append("VOID")

                    elif df['Void'].iloc[i] =="VOID":
                        void_new.append('VOID')
                        id_hold = df['Ticket ID'].iloc[i]
                    else:
                        void_new.append(0)



                df['Void'] = void_new
                df = df[df['Void'] == 0]

                df = df.drop('Void', axis=1)

                df.dropna(axis=0, how='any', inplace=True, subset=['Weight'])


                df['Ticket Date'] = df['Ticket Date'].dt.strftime('%m/%d/%Y')


                for i in range(len(df)):
                    if df['Ticket ID'].iloc[i - 1] > df['Ticket ID'].iloc[i] and df['Ticket ID'].iloc[i - 2] > df['Ticket ID'].iloc[i] :
                        split = i

                receiving_df = df.iloc[:split, :]
                shipping_df = df.iloc[split:, :]


                df = receiving_df
                mat_list = []
                mat_list_full = []
                for i in range(len(df)):
                    if df['Area'].iloc[i] not in mat_list:
                        mat_list.append(df['Area'].iloc[i])
                    mat_list_full.append(df['Area'].iloc[i])
                mat_count = []
                mat_sum = []
                for i in range(len(mat_list)):
                    mat_count.append(mat_list_full.count(mat_list[i]))
                    dfhold = df[df['Area'] == mat_list[i]]
                    mat_sum.append(dfhold['Weight'].sum())
                rec_dict = {'Mat': mat_list, 'Count': mat_count, 'Sum': mat_sum}
                rec_totals = pd.DataFrame(rec_dict)

                df = shipping_df
                mat_list = []
                mat_list_full = []
                for i in range(len(df)):
                    if df['Area'].iloc[i] not in mat_list:
                        mat_list.append(df['Area'].iloc[i])
                    mat_list_full.append(df['Area'].iloc[i])
                mat_count = []
                mat_sum = []
                for i in range(len(mat_list)):
                    mat_count.append(mat_list_full.count(mat_list[i]))
                    dfhold = df[df['Area'] == mat_list[i]]
                    mat_sum.append(dfhold['Weight'].sum())
                shp_dict = {'Mat': mat_list, 'Count': mat_count, 'Sum': mat_sum}
                shp_totals = pd.DataFrame(shp_dict)

                if '.xlsx' not in T2.get("1.0", "end-1c"):
                    string_name2 = T2.get("1.0", "end-1c") + '.xlsx'
                else:
                    string_name2 = T2.get("1.0", "end-1c")


                writer = pd.ExcelWriter(string_name2,
                                        engine='xlsxwriter')
                receiving_df.to_excel(writer, sheet_name="Receiving", index=False)
                shipping_df.to_excel(writer, sheet_name="Shipping", index=False)
                rec_totals.to_excel(writer, sheet_name="Rec_Totals", index=False)
                shp_totals.to_excel(writer, sheet_name="Shp_Totals", index=False)
                writer.save()
                tk.messagebox.showinfo(title=None, message='Complete')




#############################################################################################################
class PageOne(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

class PageTwo(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)


print("***************************************************************************")
print("Data Formatting Tools Version 1.0 (beta) Monterey Metal Works and Iron \nBrought to you by ajl130 and Texas State University")
print(" ")
print("Troubleshooting Tips: ")
print("1) Make sure sheet names are in the format 'Sheet1' and 'Sheet2'.")
print("2) 'Scale Ticket Data' is designed to handle at most 2 sheets which should cover one year worth of data.")
print("3) Make sure that the correct data type is chosen for each specific program. It is specfic and not universal. ")
print("***************************************************************************")
app = holdy_boi()
separator = Frame(height=2, bd=1, relief=SUNKEN)
separator.pack(fill=X, padx=5, pady=5)
app.geometry("300x300")
app.title('Reformat Data')
app.mainloop()


