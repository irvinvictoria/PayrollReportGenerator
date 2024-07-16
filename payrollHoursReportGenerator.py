from tkinter import *
from tkinter import filedialog
import pandas as pd
import numpy as numpy
import sys


payRates= pd.DataFrame()

# Opens the report CSV file
# Creates the payRate dictionary
# calls the generate report method

def runReport():
    inputDate = dateTextBox.get("1.0", "end-1c")
    folderPathG1 = filedialog.askdirectory()
    filepathGl = filedialog.askopenfilename()
    generateReport(filepathGl, folderPathG1,inputDate)

# Creates the add pay rate window
def openPayrateWindow():
    global payRates
    payRates = pd.read_csv('resources/payRates.csv')
    def close():
        addPayrateWindow.quit()
        addPayrateWindow.destroy()
    def addPay():
        global payRates
        newEEID = eeidText.get("1.0", "end-1c")
        newPayRate = payrateText.get("1.0", "end-1c")
        newPayrateRow = {'EEID':newEEID, 'payRate': newPayRate}
        new_payIndex = len(payRates)
        payRates.loc[new_payIndex] = newPayrateRow
        payRates.to_csv('resources/payRates.csv', index=False)
        payRates = pd.read_csv('resources/payRates.csv')
        eeidText.delete('1.0', END)
        payrateText.delete('1.0', END)
    addPayrateWindow = Tk()
    addPayrateWindow.minsize(400, 300)
    eeidLabel = Label(addPayrateWindow, text="Enter EEID")
    eeidText = Text(addPayrateWindow, height=1, width=8)
    payrateLabel = Label(addPayrateWindow, text="Enter Pay Rate")
    payrateText = Text(addPayrateWindow, height=1, width=8)
    addPayButton = Button(addPayrateWindow, text="Add Pay Rate", command=addPay)
    closeWindowButton = Button(addPayrateWindow, text="Close Window", command=close)
    eeidLabel.pack()
    eeidText.pack()
    payrateLabel.pack()
    payrateText.pack()
    addPayButton.pack()
    closeWindowButton.pack()
    addPayrateWindow.mainloop()


# Generates the final report
def generateReport(file, folderPath, dateTxt):
    global payRates
    payRates = pd.read_csv('resources/payRates.csv')
    
    #payRates = getPayRateFile()
    #selects headers from row 4
    df = pd.read_csv(file, header=4)

    missingEeid = []
    # Looks for missing EEID's
    for index, row in df.iterrows():
        if row["DIS"] == "J/C" or row['DIS'] == "W/O" or row["DIS"] == "G/L":
            eeid = int(row["EEID"]) 
            try:
                payRates.loc[payRates['EEID'] == eeid, 'payRate'].values[0]
            except:
                if not eeid in missingEeid:
                    missingEeid.append(eeid)

    # Prompts user to enter pay rates if not eeid not found
    if len(missingEeid) != 0:
        # Closes the entire program
        def closeProgram():
            sys.exit()
        # Function for addPay button. Adds payrates and eeids to the payRates.csv file
        def addPay():
            global payRates
            newEEID = eeidText1.get("1.0", "end-1c")
            newPayRate = payRateText1.get("1.0", "end-1c")
            newPayrateRow = {'EEID':newEEID, 'payRate': newPayRate}
            new_payIndex = len(payRates)
            payRates.loc[new_payIndex] = newPayrateRow
            payRates.to_csv('resources/payRates.csv', index=False)
            payRates = pd.read_csv('resources/payRates.csv')
            eeidText1.delete('1.0', END)
            payRateText1.delete('1.0', END)
        # Function for the continue button. Lets the program continue running.
        def continueProgram():
            errWindow1.quit()
            errWindow1.destroy()
        # Creates the error window
        errWindow1 = Tk()
        errWindow1.minsize(300, 200)
        errLabel1 = Label(errWindow1, text= "PayRate for "+ str(missingEeid) +" not found, please input payrates then press continue or close application")
        eeidLabel1 = Label(errWindow1, text= "Input EEID")
        eeidText1 = Text(errWindow1, height=1, width=8)
        payRateLabel1 = Label(errWindow1, text= "Input payrate")
        payRateText1 = Text(errWindow1, height=1, width=8)
        addPayRateButton1 = Button(errWindow1, text='Add Pay Rate', command=addPay)
        continueButton = Button(errWindow1, text="Continue", command=continueProgram)
        errButton1 = Button(errWindow1, text="Close Program", command=closeProgram)
        errLabel1.pack()
        eeidLabel1.pack()
        eeidText1.pack()
        payRateLabel1.pack()
        payRateText1.pack()
        addPayRateButton1.pack()
        continueButton.pack()
        errButton1.pack()
        errWindow1.mainloop()

    # Create empty dataframe and dictionary for final report
    reportDf = pd.DataFrame(columns=['Type', 'Account', 'Dept', 'Subaccount', 'Date', 'Amount', 'Hours', 'Reference', 'Description'])
    jobDistList = {}

    #inserts new headers
    df.insert(11, "Pay Rate",value=None, allow_duplicates= False)
    df.insert(12, "Reg Pay",value=None, allow_duplicates= False)
    df.insert(13, "OT Pay",value=None, allow_duplicates= False)
    df.insert(14, "DT Pay",value=None, allow_duplicates= False)
    df.insert(15, "Total Pay",value=None, allow_duplicates= False)

    #inserts pay rates, calculates: reg pay, OT pay, DT pay, and Total Pay
    for index, row in df.iterrows():
        if row["DIS"] == "J/C" or row['DIS'] == "W/O" or row["DIS"] == "G/L":
            # Finds the payrate and updates sheet
            eeid = int(row["EEID"]) 
            def getRate():
                try:
                    return payRates.loc[payRates['EEID'] == eeid, 'payRate'].values[0]
                except:
                    def closeProgram():
                        sys.exit()
                    # Creates the error window if there is a missing eeid and payrate
                    errWindow = Tk()
                    errWindow.minsize(300, 200)
                    errLabel = Label(errWindow, text= "Missing EEID's and Pay Rates. Close the app and try again.")
                    errButton = Button(errWindow, text="Close Program", command=closeProgram)
                    errLabel.pack()
                    errButton.pack()
                    errWindow.mainloop()
            rate = getRate()
            df.at[index, 'Pay Rate'] = rate
            # Calculates and updates regular pay
            regHrs = float(row["REG HR"])
            regHrsPay = round(regHrs * rate,2)
            df.at[index, 'Reg Pay'] = regHrsPay
            # Calculates and updates OT pay
            otHrs = float(row["OT HR"])
            otRate = round((rate*1.5), 2)
            otHrsPay = round(otHrs*otRate, 2)
            df.at[index, 'OT Pay'] = otHrsPay
            # Calculates DT pay
            dtHrs = float(row["DT HR"])
            dtRate = round(rate*2, 2)
            dtHrsPay = round(dtHrs*dtRate, 2)
            df.at[index, 'DT Pay'] = dtHrsPay
            # Calculates Total Pay
            totalPay = round(regHrsPay+otHrsPay+dtHrsPay, 2)
            df.at[index, 'Total Pay'] = totalPay
            # Will calculate the total per job and distribution within J/C
            if row["DIS"] == "J/C":
                # Formats the job number and division
                jobNumber = row["ALLOC ACCT"]
                division = row["ALLOC DIV"]
                # Gets rid of W from job number if there is one
                if 'W' in jobNumber:
                    jobNumber = jobNumber.replace('W','')
                else:
                    jobNumber = int(jobNumber)
                # Checks the division and formats unique job number based on division
                # Ex: 7974I for irrigation and 7974L for landscape
                if division == "IRRIGA":
                    jobNumber = str(jobNumber) + "I"
                elif division  == "LANDSC":
                    jobNumber = str(jobNumber) + "L"
                else: 
                    jobNumber = str(jobNumber) + "M"
                    
                # Adds totalPay according to jobNumber to jobDistList 
                if jobNumber in jobDistList:
                    jobDistList[jobNumber] = round(jobDistList[jobNumber] + totalPay, 2)
                else:
                    jobDistList[jobNumber] = totalPay             
           
    # Inputs info from jobDistList to empty dataframe reportDf to create final report        
    for job in jobDistList:
        amt = jobDistList[job]
        date = dateTxt
        jobNum = job[0:-1]
        jobLetter = job[-1]
        deptNum = 0
        deptName = ''
        subAcc = ''
        desc = 'Reclass Staffing Hrs'
        # Creates correct variable information for the above variables
        if jobLetter == 'I':
            deptName = 'IRRIGA'
            deptNum = 100
            jobNum = '0' + jobNum
            subAcc = 'ITL'
        elif jobLetter == 'L':
            deptName = 'LANDSC'
            deptNum = 200
            jobNum = '00' + jobNum
            subAcc = 'LTL'
        else:
            deptName = 'MAINTE'
            deptNum = 300
            jobNum = jobNum.lstrip('0')
            subAcc = 'MTL'

        # Adds rows to reportDf
        new_index = len(reportDf)
        row1 = {'Type':'J', 'Account':jobNum, 'Dept': deptName, 'Subaccount': subAcc, 'Date': date, 'Amount':amt, 'Hours':'', 'Reference':'', 'Description':desc}
        reportDf.loc[new_index] = row1
        new_index = len(reportDf)
        row2 = {'Type':'G', 'Account':'502000', 'Dept': deptNum, 'Subaccount': '', 'Date': date, 'Amount':-amt, 'Hours':'', 'Reference':'', 'Description':desc}
        reportDf.loc[new_index] = row2

    # Creates the filepath and name for the reports
    payReportPath = folderPath + '/payReport' + dateTxt + '.xlsx'
    deptReportPath = folderPath + '/deptReport' + dateTxt + '.xlsx'
    # Creates the final excel reports
    df.to_excel(payReportPath, index=False)
    reportDf.to_excel(deptReportPath,index=False)
   
# Creates main window
window = Tk()
window.minsize(400, 300)
label = Label(window, text= "Input date")
dateTextBox = Text(window, height= 1, width= 15)
buttonStart = Button(text = "Create report", command= runReport)
buttonMainAddPayrate = Button(text = "Add Pay Rate", command= openPayrateWindow)
label.pack()
dateTextBox.pack()
buttonStart.pack()
buttonMainAddPayrate.pack()
window.mainloop()

