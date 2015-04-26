__author__ = 'patdnk'

import random
import sys
import os
import datetime
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showerror

# from decimal import Decimal
from docx import Document
from os.path import abspath, expanduser

# class WeekSelector(StringVar)
#     def __init__(self):
#         self.getAllMondays()
#
#
#

# class MainFrame(Frame):
#     def __init__(self, master=NONE):
#         Frame.__init__(self, master)
#         self.master = master
#         self.grid(padx=5, pady=5)
#         self.addModules()
#
#     def addModules(self):
#         timesheet = Timesheet(self).grid(row=0,column=0)
#         invoice = Invoice(self).grid(row=0,column=1)

class Application:
    def __init__(self, parent):

        self.myParent = parent

        # define Main frame
        self.mainFrame = Frame(parent)
        self.mainFrame.grid()

        rowNumber = 0

        # define Timesheet label
        self.timesheetLabel = Label(self.mainFrame, text="Timesheet", font="System 14 bold")
        self.timesheetLabel.grid(row=rowNumber, column=0)

        # define Invoice label
        self.invoiceLabel = Label(self.mainFrame, text="Invoice", font="System 14 bold")
        self.invoiceLabel.grid(row=rowNumber, column=1)

        rowNumber += 1

        # define Timesheet frame
        self.timesheetFrame = Frame(self.mainFrame, bg="white", borderwidth=1, relief=SUNKEN, width=310, height=550)
        self.timesheetFrame.grid(row=rowNumber, column=0, padx=5, pady=5)
        self.timesheetFrame.grid_propagate(0)

        # create timesheet widgets
        self.createTimesheetWidgets()

        # define Invoice frame
        self.invoiceFrame = Frame(self.mainFrame, bg="white", borderwidth=1, relief=SUNKEN, width=310, height=550)
        self.invoiceFrame.grid(row=rowNumber, column=1, padx=5, pady=5)
        self.invoiceFrame.grid_propagate(0)

        # create invoice widgets
        self.createInvoiceFrameWidgets()

        # create helper tuples
        self.normalHoursEntryList = [self.entryMondayNH, self.entryTuesdayNH, self.entryWednesdayNH, self.entryThursdayNH, self.entryFridayNH, self.entrySaturdayNH, self.entrySundayNH]
        self.overtimeHoursEntryList = [self.entryMondayOH, self.entryTuesdayOH, self.entryWednesdayOH, self.entryThursdayOH, self.entryFridayOH, self.entrySaturdayOH, self.entrySundayOH]

        self.daysTuple = [("nh_monday",self.entryMondayNH),("nh_tuesday",self.entryTuesdayNH),("nh_wednesday",self.entryWednesdayNH),("nh_thursday",self.entryThursdayNH),("nh_friday",self.entryFridayNH),("nh_saturday",self.entrySaturdayNH),("nh_sunday",self.entrySundayNH),("oh_monday",self.entryMondayOH),("oh_tuesday",self.entryTuesdayOH),("oh_wednesday",self.entryTuesdayOH),("oh_thursday",self.entryThursdayOH),("oh_friday",self.entryFridayOH),("oh_saturday",self.entrySaturdayOH),("oh_sunday",self.entrySundayOH)];
        self.daysPlaceholders = ["monday_date", "tuesday_date", "wednesday_date", "thursday_date", "friday_date", "saturday_date", "sunday_date"]



        # # self.timesheetFrame.pack( padx=5, pady=5)
        # self.timesheetFrame.config(bg="white")
        # self.timesheetFrame.grid(padx=5, pady=5)



        # self.grid(padx=5, pady=5)
        # self.config(bg="white", borderwidth=1, relief=SUNKEN)
        # self.place(width=300, height=400)
        # self.create_widgets()




        # self.getAllMondays(2015)

    def createTimesheetWidgets(self):

        self.timesheetTemplateFilenameString = StringVar()

        # template frame
        self.timesheetTemplateFrame = Frame(self.timesheetFrame, bg="white", width=300)
        self.timesheetTemplateFrame.grid(padx=5, pady=5)

        templateRowNumber = 0

        self.timesheetTemplateFrameTitle = Label(self.timesheetTemplateFrame, text="Template", font="System 11", fg="darkgray")
        self.timesheetTemplateFrameTitle.grid(row=templateRowNumber, sticky=W, columnspan=2)
        templateRowNumber += 1

        self.selectTimesheetTemplateButton = Button(self.timesheetTemplateFrame, text = "Select", command=self.selectTimesheetTemplate)
        self.selectTimesheetTemplateButton.grid(row=templateRowNumber)

        # templateRowNumber += 1

        # self.entryTemplatePath = Entry(self.timesheetTemplateFrame, width=25)
        # self.entryTemplatePath.grid(row=templateRowNumber, column=1)
        self.templatePathString = StringVar()
        self.templatePathLabel = Label(self.timesheetTemplateFrame, textvariable=self.templatePathString, width=32, font="System 11", bg="white smoke")
        self.templatePathLabel.grid(row=templateRowNumber, column=1)


        # details frame
        self.timesheetDetailsFrame = Frame(self.timesheetFrame, bg="white", width=300)
        self.timesheetDetailsFrame.grid(padx=5, pady=5)

        detailRowNumber = 0

        self.timesheetDetailsFrameTitle = Label(self.timesheetDetailsFrame, text="Details", font="System 11", fg="darkgray")
        self.timesheetDetailsFrameTitle.grid(row=detailRowNumber, sticky=W)
        detailRowNumber += 1

        self.consultantName = Entry(self.timesheetDetailsFrame, width=35)
        self.consultantName.insert(0, 'Your name')
        self.consultantName.grid(row=detailRowNumber, sticky=W)
        detailRowNumber += 1

        self.clientName = Entry(self.timesheetDetailsFrame, width=35)
        self.clientName.insert(0, 'Client name')
        self.clientName.grid(row=detailRowNumber, sticky=W)
        detailRowNumber += 1

        # date picker frame
        self.dateFrame = Frame(self.timesheetFrame, bg="white", width=300)
        self.dateFrame.grid(padx=5, pady=5)

        dateRowNumber = 0

        self.timesheetDetailsFrameTitle = Label(self.dateFrame, text="Date", font="System 11", fg="darkgray")
        self.timesheetDetailsFrameTitle.grid(row=dateRowNumber, sticky=W)

        spaceFiller = Frame(self.dateFrame, width=258)
        spaceFiller.grid(row=dateRowNumber, column=1)

        dateRowNumber += 1

        self.dateSelected = StringVar()
        self.dateSelected.set("Select monday")

        self.dateMenu = OptionMenu(self.dateFrame, self.dateSelected, *self.getAllMondays(2015))
        # way to feed list/tuple to option menu
        # self.dateMenu = apply(OptionMenu, (self.dateFrame, self.dateSelected) + tuple(self.getAllMondays(2015)))
        self.dateMenu.grid(row=dateRowNumber, columnspan=2, sticky=W)

        # hours frame
        self.hoursFrame = Frame(self.timesheetFrame, bg="white", width=300)
        self.hoursFrame.grid(padx=5, pady=5, sticky=W)

        hoursRowNumber = 0

        self.timesheetDetailsFrameTitle = Label(self.hoursFrame, text="Hours", font="System 11", fg="darkgray")
        self.timesheetDetailsFrameTitle.grid(row=hoursRowNumber, columnspan=3, sticky=W)

        hoursRowNumber += 1

        #header labels
        self.labelDays = Label(self.hoursFrame, text="Day", font="System 12 bold")
        self.labelDays.grid(row=hoursRowNumber)
        self.labelNormalHours = Label(self.hoursFrame, text="Normal\nhours", font="System 12 bold")
        self.labelNormalHours.grid(row=hoursRowNumber, column=1)
        self.labelOvertimeHours = Label(self.hoursFrame, text="Overtime\nhours", font="System 12 bold")
        self.labelOvertimeHours.grid(row=hoursRowNumber, column=2)
        hoursRowNumber += 1

        #labels and inputs
        self.labelMonday = Label(self.hoursFrame, text="Monday", font="System 12")
        self.labelMonday.grid(row=hoursRowNumber, sticky=E)
        self.entryMondayNH = Entry(self.hoursFrame, width=2)
        self.entryMondayNH.grid(row=hoursRowNumber, column=1)
        self.entryMondayOH = Entry(self.hoursFrame, width=2)
        self.entryMondayOH.grid(row=hoursRowNumber, column=2)
        hoursRowNumber += 1

        self.labelTuesday = Label(self.hoursFrame, text="Tuesday", font="System 12")
        self.labelTuesday.grid(row=hoursRowNumber, sticky=E)
        self.entryTuesdayNH = Entry(self.hoursFrame, width=2)
        self.entryTuesdayNH.grid(row=hoursRowNumber, column=1)
        self.entryTuesdayOH = Entry(self.hoursFrame, width=2)
        self.entryTuesdayOH.grid(row=hoursRowNumber, column=2)
        hoursRowNumber += 1

        self.labelWednesday = Label(self.hoursFrame, text="Wednesday", font="System 12")
        self.labelWednesday.grid(row=hoursRowNumber, sticky=E)
        self.entryWednesdayNH = Entry(self.hoursFrame, width=2)
        self.entryWednesdayNH.grid(row=hoursRowNumber, column=1)
        self.entryWednesdayOH = Entry(self.hoursFrame, width=2)
        self.entryWednesdayOH.grid(row=hoursRowNumber, column=2)
        hoursRowNumber += 1

        self.labelThursday = Label(self.hoursFrame, text="Thursday", font="System 12")
        self.labelThursday.grid(row=hoursRowNumber, sticky=E)
        self.entryThursdayNH = Entry(self.hoursFrame, width=2)
        self.entryThursdayNH.grid(row=hoursRowNumber, column=1,)
        self.entryThursdayOH = Entry(self.hoursFrame, width=2)
        self.entryThursdayOH.grid(row=hoursRowNumber, column=2,)
        hoursRowNumber += 1

        self.labelFriday = Label(self.hoursFrame, text="Friday", font="System 12")
        self.labelFriday.grid(row=hoursRowNumber, sticky=E)
        self.entryFridayNH = Entry(self.hoursFrame,width=2)
        self.entryFridayNH.grid(row=hoursRowNumber, column=1)
        self.entryFridayOH = Entry(self.hoursFrame,width=2)
        self.entryFridayOH.grid(row=hoursRowNumber, column=2)
        hoursRowNumber += 1

        self.labelSaturday = Label(self.hoursFrame, text="Saturday", font="System 12")
        self.labelSaturday.grid(row=hoursRowNumber, sticky=E)
        self.entrySaturdayNH = Entry(self.hoursFrame, width=2)
        self.entrySaturdayNH.grid(row=hoursRowNumber, column=1)
        self.entrySaturdayOH = Entry(self.hoursFrame, width=2)
        self.entrySaturdayOH.grid(row=hoursRowNumber, column=2)
        hoursRowNumber += 1

        self.labelSunday = Label(self.hoursFrame, text="Sunday", font="System 12")
        self.labelSunday.grid(row=hoursRowNumber, sticky=E)
        self.entrySundayNH = Entry(self.hoursFrame, width=2)
        self.entrySundayNH.grid(row=hoursRowNumber, column=1)
        self.entrySundayOH = Entry(self.hoursFrame, width=2)
        self.entrySundayOH.grid(row=hoursRowNumber, column=2)
        hoursRowNumber += 1

        self.timesheetSaveDestinationPathString = StringVar()

        # template frame
        self.timesheetSavePathFrame = Frame(self.timesheetFrame, bg="white", width=300)
        self.timesheetSavePathFrame.grid(padx=5, pady=5)

        templateRowNumber = 0

        self.timesheetDestinationFrameTitle = Label(self.timesheetSavePathFrame, text="Destination", font="System 11", fg="darkgray")
        self.timesheetDestinationFrameTitle.grid(row=templateRowNumber, sticky=W, columnspan=2)
        templateRowNumber += 1

        self.selectTimesheetDestinationButton = Button(self.timesheetSavePathFrame, text = "Select", command=self.selectTimesheetDestinationFolder)
        self.selectTimesheetDestinationButton.grid(row=templateRowNumber)

        # templateRowNumber += 1

        # self.entryTemplatePath = Entry(self.timesheetTemplateFrame, width=25)
        # self.entryTemplatePath.grid(row=templateRowNumber, column=1)
        self.templateDestinationPathString = StringVar()
        self.templateDestinationPathLabel = Label(self.timesheetSavePathFrame, textvariable=self.templateDestinationPathString, width=32, font="System 11", bg="white smoke")
        self.templateDestinationPathLabel.grid(row=templateRowNumber, column=1)

        #add the creation button
        self.createButton = Button(self.timesheetFrame, text = "Create", command=self.createTimesheetWithData)
        self.createButton.grid()

    def createInvoiceFrameWidgets(self):

        self.invoiceTemplateFilenameString = StringVar()

        # template frame
        self.invoiceTemplateFrame = Frame(self.invoiceFrame, bg="white", width=300)
        self.invoiceTemplateFrame.grid(padx=5, pady=5)

        templateRowNumber = 0

        self.invoiceTemplateFrameTitle = Label(self.invoiceTemplateFrame, text="Template", font="System 11", fg="darkgray")
        self.invoiceTemplateFrameTitle.grid(row=templateRowNumber, sticky=W, columnspan=2)
        templateRowNumber += 1

        self.selectInvoiceTemplateButton = Button(self.invoiceTemplateFrame, text="Select", command=self.selectInvoiceTemplate)
        self.selectInvoiceTemplateButton.grid(row=templateRowNumber)

        # templateRowNumber += 1

        # self.entryTemplatePath = Entry(self.timesheetTemplateFrame, width=25)
        # self.entryTemplatePath.grid(row=templateRowNumber, column=1)
        self.invoicePathString = StringVar()
        self.invoicePathLabel = Label(self.invoiceTemplateFrame, textvariable=self.invoicePathString, width=32, font="System 11", bg="white smoke")
        self.invoicePathLabel.grid(row=templateRowNumber, column=1)

        # invoice details frame
        self.invoiceDetailsFrame = Frame(self.invoiceFrame, bg="white", width=300)
        self.invoiceDetailsFrame.grid(padx=5, pady=5)

        self.labelInvoiceNumber = Label(self.invoiceDetailsFrame, text="Invoice number", font="System 12")
        self.labelInvoiceNumber.grid(row=0, sticky=E)
        self.entryInvoiceNumber = Entry(self.invoiceDetailsFrame, width=5)
        self.entryInvoiceNumber.grid(row=0, column=1, sticky=W)

        self.labelInvoiceDescription = Label(self.invoiceDetailsFrame, text="Invoice description", font="System 12")
        self.labelInvoiceDescription.grid(row=1, sticky=E)
        self.entryInvoiceDescription = Entry(self.invoiceDetailsFrame, width=20)
        self.entryInvoiceDescription.grid(row=1, column=1, sticky=W)

        self.labelUnitsNumber = Label(self.invoiceDetailsFrame, text="Units number", font="System 12")
        self.labelUnitsNumber.grid(row=2, sticky=E)
        self.entryUnitsNumber = Entry(self.invoiceDetailsFrame, width=5)
        self.entryUnitsNumber.grid(row=2, column=1, sticky=W)

        self.labelUnitPrice = Label(self.invoiceDetailsFrame, text="Unit price", font="System 12")
        self.labelUnitPrice.grid(row=3, sticky=E)
        self.entryUnitPrice = Entry(self.invoiceDetailsFrame, width=5)
        self.entryUnitPrice.grid(row=3, column=1, sticky=W)

        self.timesheetSaveDestinationPathString = StringVar()

        # destination frame
        self.invoiceSavePathFrame = Frame(self.invoiceFrame, bg="white", width=300)
        self.invoiceSavePathFrame.grid(padx=5, pady=5)

        templateRowNumber = 0

        self.invoiceDestinationFrameTitle = Label(self.invoiceSavePathFrame, text="Destination", font="System 11", fg="darkgray")
        self.invoiceDestinationFrameTitle.grid(row=templateRowNumber, sticky=W, columnspan=2)
        templateRowNumber += 1

        self.selectInvoiceDestinationButton = Button(self.invoiceSavePathFrame, text = "Select", command=self.selectInvoiceDestinationFolder)
        self.selectInvoiceDestinationButton.grid(row=templateRowNumber)

        # templateRowNumber += 1

        # self.entryTemplatePath = Entry(self.timesheetTemplateFrame, width=25)
        # self.entryTemplatePath.grid(row=templateRowNumber, column=1)
        self.invoiceDestinationPathString = StringVar()
        self.invoiceDestinationPathLabel = Label(self.invoiceSavePathFrame, textvariable=self.invoiceDestinationPathString, width=32, font="System 11", bg="white smoke")
        self.invoiceDestinationPathLabel.grid(row=templateRowNumber, column=1)


    def getAllMondays(self, year):

        oneday = datetime.timedelta(days=1)
        oneweek = datetime.timedelta(days=7)

        start = datetime.date(year=year, month=1, day=1)
        while start.weekday() != 0:
            start += oneday

        self.mondays = []
        while start.year == year:
            self.mondays.append(start)
            start += oneweek

        print(self.mondays)
        return self.mondays

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

    def selectTimesheetTemplate(self):
        timesheetTemplateFilename = askopenfilename(filetypes=(("Word 2007 doc files", "*.docx"),
                                                               ("All files", "*.*")))
        if timesheetTemplateFilename:
            try:
                self.timesheetTemplateFilenameString.set(timesheetTemplateFilename)
                urlParts = timesheetTemplateFilename.rsplit("/", 2)
                # trimmedTemplatePath = ".../" + urlParts[1] + "/" + urlParts[2]
                trimmedTemplatePath = ".../" + urlParts[2]
                self.templatePathString.set(trimmedTemplatePath)
                print(trimmedTemplatePath)
            except:
                showerror("Open Source File", "Failed to read file\n'%s'" % timesheetTemplateFilename)
            return

    def selectTimesheetDestinationFolder(self):
        timesheetSavingDirectory  = askdirectory(title="Please select a directory")

        if len(timesheetSavingDirectory) > 0:
            try:
                self.timesheetSaveDestinationPathString.set(timesheetSavingDirectory)
                urlParts = timesheetSavingDirectory.rsplit("/", 2)
                trimmedTemplatePath = ".../" + urlParts[2]
                self.templateDestinationPathString.set(trimmedTemplatePath)
                print(timesheetSavingDirectory)
            except:
                showerror("Select directory", "Failed to open directory\n'%s'" % timesheetSavingDirectory)
            return

    def selectInvoiceTemplate(self):
        return

    def selectInvoiceDestinationFolder(self):
        return

    def createTimesheetWithData(self):

        # self.timesheetDocumentTemplatePath = "/users/patdynek/Documents/Maze Sys Ltd docs/templates/wa_timesheet_template.docx"
        self.timesheetDocumentTemplatePath = abspath(expanduser("~/") + 'templates/wa_timesheet_template.docx')
        self.openTimesheetDocumentTemplate()
        self.insertHoursValues(self.daysTuple)
        self.insertDatesStartingFrom(self.mondays[15])
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


    def insertDatesStartingFrom(self,mondayDate):

        weekDays = []
        self.dates = []

        #get dates for next 7 days from monday
        for i in range(0,7):
            d = mondayDate + datetime.timedelta(days=i)
            weekDays.append(d)

        #create list with tuples (placeholder, date)
        self.dates = list(zip(self.daysPlaceholders, weekDays))

        if weekDays.__len__() == 7:
            for table in self.document.tables:
                for cell in table._cells:
                        for (phDay,dayDate) in self.dates:
                            if phDay in cell.text:
                                for paragraph in cell.paragraphs:
                                    print(paragraph.text)
                                    paragraph.text = dayDate.strftime("%d/%m/%Y")

    #save timesheet document
    def saveTimeSheetDocument(self):
        self.document.save("/users/patdynek/Documents/Maze Sys Ltd docs/templates/saved_timesheet.docx")


    #todo add support for SQLite to save data when generating documents (history purpose or even safe copy)


root = Tk()
root.title("Timesheets & Invoices")
root.geometry("650x610")
root.config(padx=5, pady=5)
root.grid_propagate(0)
# timesheetModule = Timesheet(root)
# invoiceModule = Invoice(root)

app = Application(root)

root.mainloop()