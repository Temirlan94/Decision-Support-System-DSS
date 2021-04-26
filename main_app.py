
from tkinter import *
from tkinter import messagebox
from PIL import Image, ImageTk
import datetime as dt
import pandas as pd
import numpy as np
import docx2txt
import os
class Create_Window:
    # Main menu intitialization===========================================================================

    def __init__(self, window):

        self.window = window
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        #self.window.geometry("1920x1080")
        self.window.geometry(f'{screen_width}x{screen_height}')
        self.window.title("DSS_MENU")
        self.window.configure(bg="#049fd9")
        w=int(screen_width)
        h=int(screen_height)



        #self.window.attributes("-fullscreen", True)
        try:
            self.bgpic = ImageTk.PhotoImage(Image.open("pics/bgg.jpg").resize((w, h), Image.ANTIALIAS))
            self.background_label = Label(image=self.bgpic)
            self.background_label.pack()
        except FileNotFoundError:
            self.background_label=Label(bg="#049fd9")
            self.background_label.pack()

        try:
            self.susmed_logo = ImageTk.PhotoImage(Image.open("pics/susmedhouse_logo.png").resize((int(w/4.25), int(h/5.70)), Image.ANTIALIAS))
            #self.susmed_logo = ImageTk.PhotoImage(Image.open("pics/susmedhouse_logo.png").resize((600, 250), Image.ANTIALIAS))
            self.susmed_label = Label(image=self.susmed_logo)
            self.susmed_label.place(x=w/1.35, y=h/1.35)
            #self.susmed_label.place(x=1250, y=800)
        except FileNotFoundError:
            pass

        try:
            self.artecs_logo = ImageTk.PhotoImage(Image.open("pics/artecs_logo.png").resize((int(w/4.25), int(h/5.70)), Image.ANTIALIAS))
            #self.artecs_logo = ImageTk.PhotoImage(Image.open("pics/artecs_logo.png").resize((600, 250), Image.ANTIALIAS))
            self.artecs_label = Label(image=self.artecs_logo)
            self.artecs_label.place(x=w/1.35, y=h/1.80)
            #self.artecs_label.place(x=1250, y=500)
        except FileNotFoundError:
            pass

        # Date&Time intitialization===========================================================================
        self.date = Label(window, text=f"{dt.datetime.now():%a, %b %d %Y}", fg="white", bg="#049fd9", font=("helvetica", int(w/64)))
        self.date.place(x=w/1.25, y=h/16.5)
        #self.date.place(x=1350, y=50)

        self.currentTime=f"{dt.datetime.now():%H:%M %p}"
        self.time = Label(window, text=self.currentTime, fg="white", bg="#049fd9", font=("helvetica",int(w/64)))
        self.time.place(x=w/1.25, y=h/9.5)
        #self.time.place(x=1540, y=100)


        # Button creation=====================================================================================
        #self.button1= ImageTk.PhotoImage(Image.open("button_pics/button1.png").resize((int(w/10.5), int(h/10.8)), Image.ANTIALIAS))
        #self.ePrice = Button(window,image=self.button1,bg="Blue",borderwidth=0,command=self.ePriceShow)
        self.ePrice = Button(window, text="Electricity Prices", fg="blue", bg="#fceea7",width=int(w/102.4), height=int(h/288), font=int(w / 102), command=self.ePriceShow)
        #self.ePrice = Button(window, text="Electricity Prices", fg="blue", bg="#fceea7", width="25", height="5",font=25,command=self.ePriceShow)
        self.ePrice.place(x=w/50, y=h/50)
        #self.ePrice.place(x=50, y=50)

        self.cropPrices = Button(window, text="Crop Prices", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288),font=int(w/102),command=self.cropPriceShow)
        self.cropPrices.place(x=w/50, y=h/10)
        #self.cropPrices.place(x=50, y=160)
        self.cropWindow=None #not to open same window multiple times

        self.cropPrices = Button(window, text="Greenhouse conditions", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288), font=int(w/102),command=self.greenConditionShow)
        self.cropPrices.place(x=w/50, y=h/5.5)
        #self.cropPrices.place(x=50, y=270)

        self.cropPrices = Button(window, text="Crop status", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288),font=int(w/102),command=self.cropStatus)
        self.cropPrices.place(x=w/50, y=h/3.8)
        #self.cropPrices.place(x=50, y=380)

        self.cropPrices = Button(window, text="Greenhouse controls", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288), font=int(w/102),command=self.greenHouseControls)
        self.cropPrices.place(x=w/5, y=h/50)
        #self.cropPrices.place(x=500, y=50)

        self.cropPrices = Button(window, text="Feasibility report", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288), font=int(w/102),command=self.feasibilityReportShow)
        self.cropPrices.place(x=w/5, y=h/10)
        #self.cropPrices.place(x=500, y=160)

        self.cropPrices = Button(window, text="Instructions", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288),font=int(w/102),command=self.instructionShow)
        self.cropPrices.place(x=w/5, y=h/5.5)
        #self.cropPrices.place(x=500, y=270)

        self.cropPrices = Button(window, text="Warnings", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288), font=int(w/102),command=self.greenHouseWarnings)
        self.cropPrices.place(x=w/5, y=h/3.8)
        #self.cropPrices.place(x=500, y=380)

        self.close_button = Button(window, text="Close", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288),font=int(w/102), command=window.quit)
        self.close_button.place(x=w/50, y=h/1.4)
        #self.close_button.place(x=50, y=880)



    def ePriceShow(self):
            currentMonth = text = f"{dt.datetime.now():%b}"
            currentYear= int(f"{dt.datetime.now():%Y}")
            try:
                ePrice_data = pd.read_excel(r'excelFiles/elektrik.xls')
                all_data = pd.DataFrame(ePrice_data)
                all_data['Condition1'] = all_data['Month'].apply(lambda x: 'True' if x == currentMonth else 'False')
                all_data['Condition2'] = all_data['Year'].apply(lambda x: 'True' if x == currentYear else 'False')
                all_data['Condition3'] = np.where((all_data['Condition1'] == all_data['Condition2']), all_data['Condition1'], "False")
                showPrice = all_data.loc[all_data['Condition3'] == 'True']['Price'].values[0]
                messagebox.showinfo("showinfo", "{} {} Price is: {} TL".format(currentMonth,currentYear,showPrice))
                #print (all_data)
            except FileNotFoundError:
                messagebox.showinfo("showinfo", "No prices available, try again later!")



    def cropPriceShow(self):

        if self.cropWindow is None:  #not to open same window multiple times
            currentMonth = text = f"{dt.datetime.now():%b}"
            currentYear= int(f"{dt.datetime.now():%Y}")

            def selection1():
                tomatoPrice = pd.read_excel(r'excelFiles/tomato.xls')
                all_data = pd.DataFrame(tomatoPrice)
                all_data['Condition1'] = all_data['Month'].apply(lambda x: 'True' if x == currentMonth else 'False')
                all_data['Condition2'] = all_data['Year'].apply(lambda x: 'True' if x == currentYear else 'False')
                all_data['Condition3'] = np.where((all_data['Condition1'] == all_data['Condition2']),all_data['Condition1'], "False")
                showPrice = all_data.loc[all_data['Condition3'] == 'True']['Price'].values[0]
                messagebox.showinfo("showinfo", "{} {} Tomato price is: {} TL".format(currentMonth,currentYear,showPrice))
                self.cropWindow.destroy()
                self.cropWindow = None
            def selection2():
                lettucePrice = pd.read_excel(r'excelFiles/lettuce.xls')
                all_data = pd.DataFrame(lettucePrice)
                all_data['Condition1'] = all_data['Month'].apply(lambda x: 'True' if x == currentMonth else 'False')
                all_data['Condition2'] = all_data['Year'].apply(lambda x: 'True' if x == currentYear else 'False')
                all_data['Condition3'] = np.where((all_data['Condition1'] == all_data['Condition2']),all_data['Condition1'], "False")
                showPrice = all_data.loc[all_data['Condition3'] == 'True']['Price'].values[0]
                messagebox.showinfo("showinfo", "{} {} Lettuce price is: {} TL".format(currentMonth,currentYear,showPrice))
                self.cropWindow.destroy()
                self.cropWindow = None
            def selection3():
                pepperPrice = pd.read_excel(r'excelFiles/pepper.xls')
                all_data = pd.DataFrame(pepperPrice)
                all_data['Condition1'] = all_data['Month'].apply(lambda x: 'True' if x == currentMonth else 'False')
                all_data['Condition2'] = all_data['Year'].apply(lambda x: 'True' if x == currentYear else 'False')
                all_data['Condition3'] = np.where((all_data['Condition1'] == all_data['Condition2']),all_data['Condition1'], "False")
                showPrice = all_data.loc[all_data['Condition3'] == 'True']['Price'].values[0]
                messagebox.showinfo("showinfo", "{} {} Pepper price is: {} TL".format(currentMonth,currentYear,showPrice))
                self.cropWindow.destroy()
                self.cropWindow = None

            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            w = int(screen_width)
            h = int(screen_height)
            self.cropWindow = Tk()
            self.cropWindow.geometry(f'{int(w / 4.2)}x{int(h / 3.2)}')
            self.cropWindow.configure(bg="#049fd9")
            self.cropWindow.title("Crop Price")


            radio = IntVar()
            R1 = Radiobutton(self.cropWindow, text="Tomato",bg="#fceea7", width=int(w/102.4), height=int(h/288),font=int(w/102), variable=radio, value= 1,command=selection1)
            R1.pack(anchor=W)

            R2 = Radiobutton(self.cropWindow, text="Lettuce",bg="#fceea7", width=int(w/102.4), height=int(h/288),font=int(w/102), variable=radio, value= 2,command=selection2)
            R2.pack(anchor=W)

            R3 = Radiobutton(self.cropWindow, text="Pepper", bg="#fceea7", width=int(w/102.4), height=int(h/288),font=int(w/102), variable=radio, value= 3,command=selection3)
            R3.pack(anchor=W)


    def instructionShow(self):
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        w = int(screen_width)
        h = int(screen_height)
        try:
            self.instructionBox = Tk()
            self.instructionBox.geometry(f'{int(w / 2.56)}x{int(h / 1.4)}')
            labelframe1 = LabelFrame(self.instructionBox, text="Instructions")
            labelframe1.pack(fill="both", expand="yes")
            result = docx2txt.process("instructions/cleaning.docx")
            #mssg="Clean************************************************"
            toplabel = Label(labelframe1, text=result)
            toplabel.place(x=0,y=50)
        except FileNotFoundError:
            messagebox.showinfo("showinfo", "No instruction available, try again later!")


    def feasibilityReportShow(self):
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        w = int(screen_width)
        h = int(screen_height)
        try:
            self.messageBox = Tk()
            self.messageBox.geometry(f'{int(w / 2.56)}x{int(h / 1.4)}')
            ePrice_data = pd.read_excel(r'excelFiles/elektrik.xls')
            all_data = pd.DataFrame(ePrice_data)

            labelframe1 = LabelFrame(self.messageBox, text="Electricity Price")
            labelframe1.pack(fill="both", expand="yes")

            toplabel = Label(labelframe1, text=all_data)
            toplabel.place(x=0, y=h/28.8)
        except FileNotFoundError:
            messagebox.showinfo("showinfo", "No report available, try again later!")

    def greenConditionShow(self):

        def showImg(room_id):
            if room_id==1:


                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 1 Status", width=int(w / 1.5),height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 1_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 1_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w/70,y=h/60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)",font=int(w/102))
                    toplabel.place(x=w/3.5,y=h/2.5)


            elif room_id==2:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 2 Status", width=int(w / 1.5),height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)

                try:
                    self.pic1 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 2_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 2_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)

            elif room_id == 3:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 3 Status", width=int(w / 1.5),height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 3_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 3_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)

            elif room_id == 4:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 4 Status", width=int(w / 1.5),height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 4_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 4_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)

            elif room_id == 5:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 5 Status", width=int(w / 1.5), height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 5_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(Image.open("PSO_Check/Oda 5_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)

            elif room_id == 6:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 6 Status", width=int(w / 1.5), height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 6_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 6_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)

            elif room_id == 7:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 7 Status", width=int(w / 1.5),height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage(
                        Image.open("PSO_Check/Oda 7_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(
                        Image.open("PSO_Check/Oda 7_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)

            elif room_id == 8:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 8 Status", width=int(w / 1.5),height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage(
                        Image.open("PSO_Check/Oda 8_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(
                        Image.open("PSO_Check/Oda 8_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)


                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)

            elif room_id == 9:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 9 Status", width=int(w / 1.5), height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage(
                        Image.open("PSO_Check/Oda 9_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(
                        Image.open("PSO_Check/Oda 9_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)

            elif room_id == 10:
                self.labelframe = LabelFrame(self.conditionWindow, text="Cabin 10 Status", width=int(w / 1.5),height=int(h))
                self.labelframe.place(x=w / 2.6, y=h / 25)
                try:
                    self.pic1 = ImageTk.PhotoImage(
                        Image.open("PSO_Check/Oda 10_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(
                        Image.open("PSO_Check/Oda 10_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 70, y=h / 2.5)

                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)




        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        w = int(screen_width)
        h = int(screen_height)
        self.conditionWindow = Toplevel()    #Toplevel
        self.conditionWindow.geometry(f'{int(w)}x{int(h)}')
        self.conditionWindow.configure(bg="#049fd9")
        self.conditionWindow.title("Greenhouse Conditions")



        self.buttonOda = Button(self.conditionWindow, text="Cabin_1", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(1))
        self.buttonOda.place(x=w/55, y=h/25)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_2", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(2))
        self.buttonOda.place(x=w/55, y=h/8)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_3", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(3))
        self.buttonOda.place(x=w/55, y=h /4.8)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_4", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(4))
        self.buttonOda.place(x=w/55, y=h/3.4)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_5", fg="blue", bg="#fceea7", width=int(w / 102.4), height=int(h / 288), font=int(w / 102),command=lambda: showImg(5))
        self.buttonOda.place(x=w / 55, y=h / 2.65)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_6", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(6))
        self.buttonOda.place(x=w/5, y=h/25)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_7", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(7))
        self.buttonOda.place(x=w/5, y=h/8)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_8", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(8))
        self.buttonOda.place(x=w/5, y=h /4.8)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_9", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(9))
        self.buttonOda.place(x=w / 5, y=h / 3.4)

        self.buttonOda = Button(self.conditionWindow, text="Cabin_10", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: showImg(10))
        self.buttonOda.place(x=w / 5, y=h / 2.65)


        self.close_button = Button(self.conditionWindow, text="Close", fg="blue", bg="#fceea7", width=int(w/102.4), height=int(h/288),font=int(w/102), command=self.conditionWindow.destroy)
        self.close_button.place(x=w/50, y=h/1.4)


    def cropStatus(self):

        def show_status(plant_id):
            if plant_id==1:
                self.labelframe = LabelFrame(self.statusWindow, text="Tomato", width=int(w / 1.2),height=int(h))
                self.labelframe.place(x=w / 4.6, y=h / 25)
            elif plant_id==2:
                self.labelframe = LabelFrame(self.statusWindow, text="Lettuce", width=int(w / 1.2), height=int(h))
                self.labelframe.place(x=w / 4.6, y=h / 25)

                try:
                    self.pic1 = ImageTk.PhotoImage(Image.open("lettuceWeek/samples/orig.jpg").resize((int(w / 8), int(h /6)),Image.ANTIALIAS))
                    self.pic2 = ImageTk.PhotoImage(Image.open("lettuceWeek/samples/fg1.jpg").resize((int(w / 8), int(h / 6)), Image.ANTIALIAS))
                    self.pic3 = ImageTk.PhotoImage(Image.open("lettuceWeek/samples/fg2.jpg").resize((int(w / 8), int(h / 6)), Image.ANTIALIAS))


                    toplabel = Label(self.labelframe, image=self.pic1)
                    toplabel.place(x=w / 70, y=h / 4)

                    toplabel1 = Label(self.labelframe, image=self.pic2)
                    toplabel1.place(x=w / 6, y=h / 4)

                    toplabel1 = Label(self.labelframe, image=self.pic3)
                    toplabel1.place(x=w / 3.1, y=h / 4)

                    self.week1 = ImageTk.PhotoImage(Image.open("lettuceWeek/artecs_lettuce/1Week.jpg").resize((int(w / 10), int(h / 10)),Image.ANTIALIAS))
                    self.week2 = ImageTk.PhotoImage(Image.open("lettuceWeek/artecs_lettuce/2Week.jpg").resize((int(w / 10), int(h / 10)),Image.ANTIALIAS))
                    self.week3 = ImageTk.PhotoImage(Image.open("lettuceWeek/artecs_lettuce/3Week.jpg").resize((int(w / 10), int(h / 10)),Image.ANTIALIAS))
                    self.week4 = ImageTk.PhotoImage(Image.open("lettuceWeek/artecs_lettuce/4Week.jpg").resize((int(w / 10), int(h / 10)),Image.ANTIALIAS))
                    self.week5 = ImageTk.PhotoImage(Image.open("lettuceWeek/artecs_lettuce/5Week.jpg").resize((int(w / 10), int(h / 10)),Image.ANTIALIAS))
                    self.week6 = ImageTk.PhotoImage(Image.open("lettuceWeek/artecs_lettuce/6Week.jpg").resize((int(w / 10), int(h / 10)),Image.ANTIALIAS))
                    self.week7 = ImageTk.PhotoImage(Image.open("lettuceWeek/artecs_lettuce/7Week.jpg").resize((int(w / 10), int(h / 10)),Image.ANTIALIAS))

                    toplabel = Label(self.labelframe, image=self.week1)
                    toplabel.place(x=w / 70, y=h / 60)
                    toplabel = Label(self.labelframe, text="1 Week", font=int(w / 102))
                    toplabel.place(x=w / 70, y=h / 60)

                    toplabel = Label(self.labelframe, image=self.week2)
                    toplabel.place(x=w / 10, y=h / 60)
                    toplabel = Label(self.labelframe, text="2 Week", font=int(w / 102))
                    toplabel.place(x=w / 10, y=h / 60)

                    toplabel = Label(self.labelframe, image=self.week3)
                    toplabel.place(x=w / 5, y=h / 60)
                    toplabel = Label(self.labelframe, text="3 Week", font=int(w / 102))
                    toplabel.place(x=w / 5, y=h / 60)

                    toplabel = Label(self.labelframe, image=self.week4)
                    toplabel.place(x=w / 3.5, y=h / 60)
                    toplabel = Label(self.labelframe, text="4 Week", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 60)

                    toplabel = Label(self.labelframe, image=self.week5)
                    toplabel.place(x=w / 2.6, y=h / 60)
                    toplabel = Label(self.labelframe, text="5 Week", font=int(w / 102))
                    toplabel.place(x=w / 2.6, y=h / 60)

                    toplabel = Label(self.labelframe, image=self.week6)
                    toplabel.place(x=w / 2.1, y=h / 60)
                    toplabel = Label(self.labelframe, text="6 Week", font=int(w / 102))
                    toplabel.place(x=w / 2.1, y=h / 60)

                    toplabel = Label(self.labelframe, image=self.week7)
                    toplabel.place(x=w / 1.8, y=h / 60)
                    toplabel = Label(self.labelframe, text="7 Week", font=int(w / 102))
                    toplabel.place(x=w / 1.8, y=h / 60)


                except FileNotFoundError:
                    toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                    toplabel.place(x=w / 3.5, y=h / 2.5)


                try:
                    lsize = pd.read_excel(r'excelFiles/size.xls')
                    all_data = pd.DataFrame(lsize)
                    lettuceW=lsize["Width"].head(1)
                    lettuceH=lsize["Height"].head(1)
                    lettuceA=int(lettuceW*lettuceH)
                    #print(lettuceA)
                    if lettuceA>50 and lettuceA <=54:
                        toplabel = Label(self.labelframe, text="1 Week", font=int(w / 102))
                        toplabel.place(x=w / 70, y=h / 2.2)
                    elif lettuceA>54 and lettuceA<=90:
                        toplabel = Label(self.labelframe, text="2 Week", font=int(w / 102))
                        toplabel.place(x=w / 70, y=h / 2.2)
                    elif lettuceA>90 and lettuceA<=143:
                        toplabel = Label(self.labelframe, text="3 Week", font=int(w / 102))
                        toplabel.place(x=w / 70, y=h / 2.2)
                    elif lettuceA>143 and lettuceA<=182:
                        toplabel = Label(self.labelframe, text="4 Week", font=int(w / 102))
                        toplabel.place(x=w / 70, y=h / 2.2)
                    elif lettuceA > 182 and lettuceA <= 270:
                        toplabel = Label(self.labelframe, text="5 Week", font=int(w / 102))
                        toplabel.place(x=w / 70, y=h / 2.2)
                    elif lettuceA>270 and lettuceA<=399:
                        toplabel = Label(self.labelframe, text="6 Week", font=int(w / 102))
                        toplabel.place(x=w / 70, y=h / 2.2)
                    elif lettuceA>399 and lettuceA<=525:
                        toplabel = Label(self.labelframe, text="7 Week", font=int(w / 102))
                        toplabel.place(x=w / 70, y=h / 2.2)
                    else:
                        toplabel = Label(self.labelframe, text="Uknown", font=int(w / 102))
                        toplabel.place(x=w / 70, y=h / 5)
                except FileNotFoundError:
                        toplabel = Label(self.labelframe, text="LOADING... (Press Again Later!)", font=int(w / 102))
                        toplabel.place(x=w / 3.5, y=h / 2.5)




            elif plant_id==3:
                self.labelframe = LabelFrame(self.statusWindow, text="Pepper", width=int(w / 1.2), height=int(h))
                self.labelframe.place(x=w / 4.6, y=h / 25)




        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        w = int(screen_width)
        h = int(screen_height)
        self.statusWindow = Toplevel()  # Toplevel
        self.statusWindow.geometry(f'{int(w)}x{int(h)}')
        self.statusWindow.configure(bg="#049fd9")
        self.statusWindow.title("Crop Status")




        self.buttonOda = Button(self.statusWindow, text="Tomato", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102),command=lambda: show_status(1))
        self.buttonOda.place(x=w / 55, y=h / 25)

        self.buttonOda = Button(self.statusWindow, text="Lettuce", fg="blue", bg="#fceea7", width=int(w / 102.4), height=int(h / 288), font=int(w / 102),command=lambda: show_status(2))
        self.buttonOda.place(x=w / 55, y=h / 8)

        self.buttonOda = Button(self.statusWindow, text="Pepper", fg="blue", bg="#fceea7", width=int(w / 102.4), height=int(h / 288), font=int(w / 102),command=lambda: show_status(3))
        self.buttonOda.place(x=w / 55, y=h / 4.8)

        self.close_button = Button(self.statusWindow, text="Close", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102), command=self.statusWindow.destroy)
        self.close_button.place(x=w / 50, y=h / 1.4)

    def greenHouseControls(self):
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        w = int(screen_width)
        h = int(screen_height)
        self.controlWindow = Toplevel()  # Toplevel
        self.controlWindow.geometry(f'{int(w)}x{int(h)}')
        self.controlWindow.configure(bg="#049fd9")
        self.controlWindow.title("Greenhouse Controls")

        self.close_button = Button(self.controlWindow, text="Close", fg="blue", bg="#fceea7", width=int(w / 102.4), height=int(h / 288), font=int(w / 102), command=self.controlWindow.destroy)
        self.close_button.place(x=w / 50, y=h / 1.4)


    def greenHouseWarnings(self):

        def dataCheck():

            #Checking if Price data is available
            try:
                ePrice_data = pd.read_excel(r'excelFiles/elektrik.xls')
                ePrice=1
            except FileNotFoundError:
                ePrice=0

            try:
                lettucePrice = pd.read_excel(r'excelFiles/lettuce.xls')
                lPrice=1
            except FileNotFoundError:
                lPrice=0

            try:
                tomatoPrice = pd.read_excel(r'excelFiles/tomato.xls')
                tPrice=1
            except FileNotFoundError:
                tPrice=0

            try:
                pepperPrice = pd.read_excel(r'excelFiles/pepper.xls')
                pPrice=1
            except FileNotFoundError:
                pPrice=0

            try:
                result = docx2txt.process("instructions/cleaning.docx")
                instruction=1
            except FileNotFoundError:
                instruction=0

            # Checking if Cabin data is available
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 1_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 1_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room1=1
            except FileNotFoundError:
                room1=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 2_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 2_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room2=1
            except FileNotFoundError:
                room2=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 3_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 3_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room3=1
            except FileNotFoundError:
                room3=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 4_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 4_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room4=1
            except FileNotFoundError:
                room4=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 5_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 5_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room5=1
            except FileNotFoundError:
                room5=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 6_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 6_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room6=1
            except FileNotFoundError:
                room6=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 7_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 7_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room7=1
            except FileNotFoundError:
                room7=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 8_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 8_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room8=1
            except FileNotFoundError:
                room8=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 9_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 9_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room9=1
            except FileNotFoundError:
                room9=0
            try:
                self.pic1 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 10_0.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                self.pic2 = ImageTk.PhotoImage( Image.open("PSO_Check/Oda 10_1.png").resize((int(w / 2.9), int(h / 2.9)), Image.ANTIALIAS))
                room10=1
            except FileNotFoundError:
                room10=0



            self.labelframe = LabelFrame(self.warningWindow, text="Data Check", width=int(w / 1.2), height=int(h))
            self.labelframe.place(x=w / 4.6, y=h / 25)

            if ePrice==1:
                toplabel = Label(self.labelframe, text="Electricity Price------>OK")
                toplabel.place(x=w / 70, y=h / 60)
            else:
                toplabel = Label(self.labelframe, text="Electricity Price------>Not Found")
                toplabel.place(x=w / 70, y=h / 60)

            if lPrice==1:
                toplabel = Label(self.labelframe, text="Lettuce Price------>OK")
                toplabel.place(x=w / 70, y=h / 30)
            else:
                toplabel = Label(self.labelframe, text="Lettuce Price------>Not Found")
                toplabel.place(x=w / 70, y=h / 30)

            if tPrice == 1:
                toplabel = Label(self.labelframe, text="Tomato Price------>OK")
                toplabel.place(x=w / 70, y=h / 20)
            else:
                toplabel = Label(self.labelframe, text="Tomato Price------>Not Found")
                toplabel.place(x=w / 70, y=h / 20)

            if pPrice == 1:
                toplabel = Label(self.labelframe, text="Pepper Price------>OK")
                toplabel.place(x=w / 70, y=h / 15)
            else:
                toplabel = Label(self.labelframe, text="Pepper Price------>Not Found")
                toplabel.place(x=w / 70, y=h / 15)

            if instruction == 1:
                toplabel = Label(self.labelframe, text="Instruction------>OK")
                toplabel.place(x=w / 70, y=h / 12.2)
            else:
                toplabel = Label(self.labelframe, text="Instruction------>Not Found")
                toplabel.place(x=w / 70, y=h / 12.2)

            if room1 == 1:
                toplabel = Label(self.labelframe, text="Cabin 1------>OK")
                toplabel.place(x=w / 70, y=h / 10.5)
            else:
                toplabel = Label(self.labelframe, text="Cabin 1----->Not Found")
                toplabel.place(x=w / 70, y=h / 10.5)

            if room2 == 1:
                toplabel = Label(self.labelframe, text="Cabin 2------>OK")
                toplabel.place(x=w / 70, y=h / 9)
            else:
                toplabel = Label(self.labelframe, text="Cabin 2------>Not Found")
                toplabel.place(x=w / 70, y=h / 9)

            if room3 == 1:
                toplabel = Label(self.labelframe, text="Cabin 3------>OK")
                toplabel.place(x=w / 70, y=h / 8)
            else:
                toplabel = Label(self.labelframe, text="Cabin 3------>Not Found")
                toplabel.place(x=w / 70, y=h / 8)

            if room4 == 1:
                toplabel = Label(self.labelframe, text="Cabin 4------>OK")
                toplabel.place(x=w / 70, y=h / 7.2)
            else:
                toplabel = Label(self.labelframe, text="Cabin 4------>Not Found")
                toplabel.place(x=w / 70, y=h / 7.2)

            if room5 == 1:
                toplabel = Label(self.labelframe, text="Cabin 5------>OK")
                toplabel.place(x=w / 70, y=h / 6.5)
            else:
                toplabel = Label(self.labelframe, text="Cabin 5------>Not Found")
                toplabel.place(x=w / 70, y=h / 6.5)

            if room6 == 1:
                toplabel = Label(self.labelframe, text="Cabin 6------>OK")
                toplabel.place(x=w / 70, y=h / 5.9)
            else:
                toplabel = Label(self.labelframe, text="Cabin 6------>Not Found")
                toplabel.place(x=w / 70, y=h / 5.9)

            if room7 == 1:
                toplabel = Label(self.labelframe, text="Cabin 7------>OK")
                toplabel.place(x=w / 70, y=h / 5.5)
            else:
                toplabel = Label(self.labelframe, text="Cabin 7------>Not Found")
                toplabel.place(x=w / 70, y=h / 5.5)

            if room8 == 1:
                toplabel = Label(self.labelframe, text="Cabin 8------>OK")
                toplabel.place(x=w / 70, y=h / 5.1)
            else:
                toplabel = Label(self.labelframe, text="Cabin 8------>Not Found")
                toplabel.place(x=w / 70, y=h / 5.1)

            if room9 == 1:
                toplabel = Label(self.labelframe, text="Cabin 9------>OK")
                toplabel.place(x=w / 70, y=h / 4.8)
            else:
                toplabel = Label(self.labelframe, text="Cabin 9------>Not Found")
                toplabel.place(x=w / 70, y=h / 4.8)

            if room10 == 1:
                toplabel = Label(self.labelframe, text="Cabin 10------>OK")
                toplabel.place(x=w / 70, y=h / 4.5)
            else:
                toplabel = Label(self.labelframe, text="Cabin 10------>Not Found")
                toplabel.place(x=w / 70, y=h / 4.5)

            toplabel = Label(self.labelframe,text="If Not Found contact administrator for data check")
            toplabel.place(x=w / 5, y=h / 60)

        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        w = int(screen_width)
        h = int(screen_height)

        self.warningWindow = Toplevel()  # Toplevel
        self.warningWindow.geometry(f'{int(w)}x{int(h)}')
        self.warningWindow.configure(bg="#049fd9")
        self.warningWindow.title("Warnings")

        self.buttonOda = Button(self.warningWindow, text="History Logs", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102))
        self.buttonOda.place(x=w / 55, y=h / 25)

        self.buttonOda = Button(self.warningWindow, text="Data Check", fg="blue", bg="#fceea7", width=int(w / 102.4),height=int(h / 288), font=int(w / 102), command=dataCheck)
        self.buttonOda.place(x=w / 55, y=h / 8)

        self.close_button = Button(self.warningWindow, text="Close", fg="blue", bg="#fceea7", width=int(w / 102.4), height=int(h / 288), font=int(w / 102), command=self.warningWindow.destroy)
        self.close_button.place(x=w / 50, y=h / 1.4)




def main():

    #os.startfile(r"PSOgraph.exe")
    app = Tk()
    menu = Create_Window(app)
    app.mainloop()


if __name__ == "__main__":
    main()