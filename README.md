This is a python app that uses the Tkinter GUI to process csv files taken from a weekly Jonas report and creates 2 excel reports.
This was made due to there being mulitple reports having to be made by the accounting dept. With each report taking 15 minutes each and on average there are 4-5 reports that are created weekly.

You must create an executable using pyinstaller.
Within the location of where the executable is stored, you must have a folder named resources that contains a payRates.csv file that has the employee id and their pay rate, without the dollar sign(xx.xx).

Terminal Line to create executable:
pyinstaller payrollHoursReportGenerator.py --onefile --noconsole


First Report:
The first report is formatted in a way that can be imported into Jonas Construction Software. This report calculates the total amount of labor accumuluated during a week per job. 

Second Report:
The second report is the result of applying the pay rate to the hours spend from the csv file. This serves as a way to double check the work and view the amount paid to any worker per job.

The csv file included shows how the information is formatted from Jonas. Due to security reasons some of the information has been changed to keep information private.

