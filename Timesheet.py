__author__ = 'patdnk'

import random
import sys
import os
import datetime
from tkinter import *
# from tkFileDialog import askopenfilename
# from decimal import Decimal
from docx import Document
from os.path import abspath, expanduser

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

        self.normalHoursEntryList = [self.entryMondayNH, self.entryTuesdayNH, self.entryWednesdayNH, self.entryThursdayNH, self.entryFridayNH, self.entrySaturdayNH, self.entrySundayNH]
        self.overtimeHoursEntryList = [self.entryMondayOH, self.entryTuesdayOH, self.entryWednesdayOH, self.entryThursdayOH, self.entryFridayOH, self.entrySaturdayOH, self.entrySundayOH]

        self.daysTuple = [("nh_monday",self.entryMondayNH),("nh_tuesday",self.entryTuesdayNH),("nh_wednesday",self.entryWednesdayNH),("nh_thursday",self.entryThursdayNH),("nh_friday",self.entryFridayNH),("nh_saturday",self.entrySaturdayNH),("nh_sunday",self.entrySundayNH),("oh_monday",self.entryMondayOH),("oh_tuesday",self.entryTuesdayOH),("oh_wednesday",self.entryTuesdayOH),("oh_thursday",self.entryThursdayOH),("oh_friday",self.entryFridayOH),("oh_saturday",self.entrySaturdayOH),("oh_sunday",self.entrySundayOH)];


        #date selection spinbox

        # listboxDates = []
        # listboxDates = self.getAllMondays(2015)
        #
        # self.spinboxDate = Listbox(self)
        # self.spinboxDate.grid()

        self.hoursButton = Button(self, text = "Get hours", command=self.getAllHoursTotals)
        self.hoursButton.grid(row=8)

        self.createButton = Button(self, text = "Create", command=self.createTimesheetWithData)
        self.createButton.grid(row=9)

        # self.createTimesheetWithData()
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

    def getAllHoursTotals(self):

        # normal hours total calculation with empty values check
        self.normalWeekHours = float(0)
        for entry in self.normalHoursEntryList:
            if entry.get():
                normalWeekHours += float(entry.get())

        # overtime hours total calculation with empty values check
        self.overtimeWeekHours = float(0)
        for entry in self.overtimeHoursEntryList:
            if entry.get():
                overtimeWeekHours += float(entry.get())

        print("Normal hours total:", normalWeekHours)
        print("Overtime hours total:", overtimeWeekHours)
        # print(overtimeWeekHours)
        # print(self.entryMondayNH.get())
        # print("Button pressed")


    #todo create dropdown with filesystem access to get the timesheet document template

    def createTimesheetWithData(self):

        # self.timesheetDocumentTemplatePath = "/users/patdynek/Documents/Maze Sys Ltd docs/templates/wa_timesheet_template.docx"
        self.timesheetDocumentTemplatePath = abspath(expanduser("~/") + 'templates/wa_timesheet_template.docx')
        self.openTimesheetDocumentTemplate()
        self.insertHoursValues(self.daysTuple)
        self.saveTimeSheetDocument()

    #open timesheet document
    def openTimesheetDocumentTemplate(self):

        self.document = Document(self.timesheetDocumentTemplatePath)

    #insert hours values
    def insertHoursValues(self,daysTuple):

        for table in self.document.tables:
            for cell in table._cells:

                for (phDay,entryDay) in daysTuple:
                    if phDay in cell.text:
                        for paragraph in cell.paragraphs:
                            print(paragraph.text)
                            if entryDay.get():
                                paragraph.text = entryDay.get()
                            else:
                                paragraph.text = "-"

                # if 'nh_monday' in cell.text:
                #     for paragraph in cell.paragraphs:
                #         print(paragraph.text)
                #         if self.entryMondayNH.get():
                #             paragraph.text = self.entryMondayNH.get()
                #         else:
                #             paragraph.text = "-"


    #insert dates -> week days from selected monday
    # def insertDates(self):
        #

    #save timesheet document
    def saveTimeSheetDocument(self):
        self.document.save("/users/patdynek/Documents/Maze Sys Ltd docs/templates/saved_timesheet.docx")


    #todo create entry for hourly/day rate
    #todo create dropdown for selection of rate period
    #todo create dropdown with filesystem access to get the invoice document template

    # #open invoice document
    # def openInvoiceDocumentTemplate(self):
    #     #
    #
    # #insert invoice values (invoice number, invoice date, invoice period, invoice description, units, unit value, total, subtotal, vat, gross total)
    # def insertInvoiceNumber(self):
    #     #
    #
    # def insertInvoiceDate(self):
    #     #
    #
    # def insertInvoicePeriodAndDescription(self):
    #     #
    #
    # def insertInvoiceUnitValues(self):
    #     #
    #
    # def insertTotals(self):
    #     #

    #todo add support for SQLite to save data when generating documents (history purpose or even safe copy)


root = Tk()
root.title("Timesheet App")
root.geometry("600x600")
app = Application(root)
root.mainloop()