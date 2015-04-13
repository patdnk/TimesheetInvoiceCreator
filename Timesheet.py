__author__ = 'patdynek'
import random
import sys
import os
import datetime
from tkinter import *

# class WeekSelector(StringVar)
#     def __init__(self):
#         self.getAllMondays()
#
#
#

class Application(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        self.grid()
        self.create_widgets()
        self.config(bg="lightgrey")
        self.place(width=300, height=300)
        # self.getAllMondays(2015)

    def create_widgets(self):

        #todo: define string vars
        #todo: create data picker to find out week starting day
        #todo: implement docx library
        #todo: create outputs of documents
        #todo: try to write invoice module based on invoice template

        #header labels
        self.labelDays = Label(self, text = "Day", font = "System 12 bold", bg="lightgray")
        self.labelDays.grid(row=0)
        self.labelNormalHours = Label(self, text = "Normal\nhours", font = "System 12 bold", bg="lightgray")
        self.labelNormalHours.grid(row=0, column=1)
        self.labelOvertimeHours = Label(self, text = "Overtime\nhours", font = "System 12 bold", bg="lightgray")
        self.labelOvertimeHours.grid(row=0, column=2)

        #labels and inputs
        self.labelMonday = Label(self, text = "Monday", font = "System 12", bg="lightgray")
        self.labelMonday.grid(row=1, sticky=E)
        self.entryMondayNH = Entry(self,width=2)
        self.entryMondayNH.grid(row=1, column=1)
        self.entryMondayOH = Entry(self,width=2)
        self.entryMondayOH.grid(row=1, column=2)

        self.labelTuesday = Label(self, text = "Tuesday", font = "System 12", bg="lightgray")
        self.labelTuesday.grid(row=2, sticky=E)
        self.entryTuesdayNH = Entry(self,width=2)
        self.entryTuesdayNH.grid(row=2, column=1)
        self.entryTuesdayOH = Entry(self,width=2)
        self.entryTuesdayOH.grid(row=2, column=2)

        self.labelWednesday = Label(self, text = "Wednesday", font = "System 12", bg="lightgray")
        self.labelWednesday.grid(row=3, sticky=E)
        self.entryWednesdayNH = Entry(self,width=2)
        self.entryWednesdayNH.grid(row=3, column=1)
        self.entryWednesdayOH = Entry(self,width=2)
        self.entryWednesdayOH.grid(row=3, column=2)

        self.labelThursday = Label(self, text = "Thursday", font = "System 12", bg="lightgray")
        self.labelThursday.grid(row=4, sticky=E)
        self.entryThursdayNH = Entry(self,width=2)
        self.entryThursdayNH.grid(row=4, column=1,)
        self.entryThursdayOH = Entry(self,width=2)
        self.entryThursdayOH.grid(row=4, column=2,)

        self.labelFriday = Label(self, text = "Friday", font = "System 12", bg="lightgray")
        self.labelFriday.grid(row=5, sticky=E)
        self.entryFridayNH = Entry(self,width=2)
        self.entryFridayNH.grid(row=5, column=1)
        self.entryFridayOH = Entry(self,width=2)
        self.entryFridayOH.grid(row=5, column=2)

        self.labelSaturday = Label(self, text = "Saturday", font = "System 12", bg="lightgray")
        self.labelSaturday.grid(row=6, sticky=E)
        self.entrySaturdayNH = Entry(self,width=2)
        self.entrySaturdayNH.grid(row=6, column=1)
        self.entrySaturdayOH = Entry(self,width=2)
        self.entrySaturdayOH.grid(row=6, column=2)

        self.labelSunday = Label(self, text = "Sunday", font = "System 12", bg="lightgray")
        self.labelSunday.grid(row=7, sticky=E)
        self.entrySundayNH = Entry(self, width=2)
        self.entrySundayNH.grid(row=7, column=1)
        self.entrySundayOH = Entry(self, width=2)
        self.entrySundayOH.grid(row=7, column=2)


        #date selection spinbox

        # listboxDates = []
        # listboxDates = self.getAllMondays(2015)
        #
        # self.spinboxDate = Listbox(self)
        # self.spinboxDate.grid()

        self.bttn2 = Button(self, text = "Get hours", command=self.getAllHoursData())
        self.bttn2.grid(row=8)
        #
        # self.bttn3 = Button(self)
        # self.bttn3.grid()
        # self.bttn3["text"] = "Apply"

    def getAllMondays(self, year):

        oneday = datetime.timedelta(days=1)
        oneweek = datetime.timedelta(days=7)

        start = datetime.date(year=year, month=1, day=1)
        while start.weekday() != 0:
            start += oneday

        days = []
        while start.year == year:
            days.append(start)
            start += oneweek

        print(days)
        return days

    def getAllHoursData(self):
        # normalWeekHours = int(self.entryMondayNH.get()) + int(self.entryTuesdayNH.get() + int(self.entryWednesdayNH.get()) + self.entryThursdayNH.get() + self.entryFridayNH.get() + self.entrySaturdayNH.get() + self.entrySundayNH.get()
        # overtimeWeekHours = self.entryMondayOH.get() + self.entryTuesdayOH.get() + self.entryWednesdayOH.get() + self.entryThursdayOH.get() + self.entryFridayOH.get() + self.entrySaturdayOH.get() + self.entrySundayNH.get()
        # print(normalWeekHours)
        # print(overtimeWeekHours)
        print(self.entryMondayNH.get())
        print("Button pressed")



root = Tk()
root.title("Timesheet App")
root.geometry("600x600")
app = Application(root)
root.mainloop()