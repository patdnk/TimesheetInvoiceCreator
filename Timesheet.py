__author__ = 'patdynek'
import random
import sys
import os
import datetime
from tkinter import *

class Application(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        self.grid()
        self.create_widgets()

    def create_widgets(self):
        #header labels
        self.labelDays = Label(self, text = "Day", font = "System 12 bold")
        self.labelDays.grid(row=0)
        self.labelNormalHours = Label(self, text = "Normal\nhours", font = "System 12 bold")
        self.labelNormalHours.grid(row=0, column=1)
        self.labelOvertimeHours = Label(self, text = "Overtime\nhours", font = "System 12 bold")
        self.labelOvertimeHours.grid(row=0, column=2)

        self.labelMonday = Label(self, text = "Monday", font = "System 12")
        self.labelMonday.grid(row=1, sticky=E)
        self.entryMondayNH = Entry(self,width=2)
        self.entryMondayNH.grid(row=1, column=1)
        self.entryMondayOH = Entry(self,width=2)
        self.entryMondayOH.grid(row=1, column=2)

        self.labelTuesday = Label(self, text = "Tuesday", font = "System 12")
        self.labelTuesday.grid(row=2, sticky=E)
        self.entryTuesdayNH = Entry(self,width=2)
        self.entryTuesdayNH.grid(row=2, column=1)
        self.entryTuesdayOH = Entry(self,width=2)
        self.entryTuesdayOH.grid(row=2, column=2)

        self.labelWednesday = Label(self, text = "Wednesday", font = "System 12")
        self.labelWednesday.grid(row=3, sticky=E)
        self.entryWednesdayNH = Entry(self,width=2)
        self.entryWednesdayNH.grid(row=3, column=1)
        self.entryWednesdayOH = Entry(self,width=2)
        self.entryWednesdayOH.grid(row=3, column=2)

        self.labelThursday = Label(self, text = "Thursday", font = "System 12")
        self.labelThursday.grid(row=4, sticky=E)
        self.entryThursdayNH = Entry(self,width=2)
        self.entryThursdayNH.grid(row=4, column=1,)
        self.entryThursdayOH = Entry(self,width=2)
        self.entryThursdayOH.grid(row=4, column=2,)

        self.labelFriday = Label(self, text = "Friday", font = "System 12")
        self.labelFriday.grid(row=5, sticky=E)
        self.entryFridayNH = Entry(self,width=2)
        self.entryFridayNH.grid(row=5, column=1)
        self.entryFridayOH = Entry(self,width=2)
        self.entryFridayOH.grid(row=5, column=2)

        self.labelSaturday = Label(self, text = "Saturday", font = "System 12")
        self.labelSaturday.grid(row=6, sticky=E)
        self.labelSaturdayNH = Entry(self,width=2)
        self.labelSaturdayNH.grid(row=6, column=1)
        self.labelSaturdayOH = Entry(self,width=2)
        self.labelSaturdayOH.grid(row=6, column=2)

        self.labelSunday = Label(self, text = "Sunday", font = "System 12")
        self.labelSunday.grid(row=7, sticky=E)
        self.labelSundayNH = Entry(self, width=2)
        self.labelSundayNH.grid(row=7, column=1)
        self.labelSundayOH = Entry(self, width=2)
        self.labelSundayOH.grid(row=7, column=2)


        # self.bttn2 = Button(self)
        # self.bttn2.grid()
        # self.bttn2.configure(text = "Cancel")
        #
        # self.bttn3 = Button(self)
        # self.bttn3.grid()
        # self.bttn3["text"] = "Apply"

root = Tk()
root.title("Timesheet App")
root.geometry("600x600")
# mainFrame = Frame(root, width=50, height=50).pack()
app = Application(root)
root.mainloop()