import tkinter
import sys
from tkinter import *
#sys.path.append('C:/ASRAWAT/test/BasicHealthReport')
#sys.path.append('C:/ASRAWAT/test/EnablerSummaryGenerator')

sys.path.append('../VerificationReportGeneratorNewLabels')
sys.path.append('../VerificationReportDocGenerator')
#import FileOperationCommentsUpdate as ES
import BasicHealthReportGenerator as BHR
import BhrReportOperations as BRO
#import VerificationReportGenerator as VRG
#import VerificationReportDocxFormat_Generator as VRG_DOC


def retrieve_input(textBox, messageWindow, teamOrMe):
    inputValue=textBox.get("1.0","end-1c")
    print(inputValue)
    if(teamOrMe == 0):
        ES.sendProgressToMeOnly(inputValue)
    else:
        ES.sendProgressToTeam(inputValue)
    messageWindow.destroy()

def readMailcontent():
    messageWindow = tkinter.Tk()
    textBox = Text(messageWindow, height=5, width=80)
    textBox.pack()
    buttonCommit = Button(messageWindow, height=1, width=20, text="Send mail",
                          command=lambda: retrieve_input(textBox, messageWindow, 0))
    # command=lambda: retrieve_input() >>> just means do this when i press the button
    buttonCommit.pack()
    messageWindow.mainloop()

def readMailcontentForTeam():
    messageWindow = tkinter.Tk()
    textBox = Text(messageWindow, height=9, width=80)
    textBox.pack()
    buttonCommit = Button(messageWindow, height=1, width=20, text="Send mail",
                          command=lambda: retrieve_input(textBox, messageWindow, 1))
    # command=lambda: retrieve_input() >>> just means do this when i press the button
    buttonCommit.pack()
    messageWindow.mainloop()

window = tkinter.Tk()
window.title("System Reports Generator")
window.geometry("750x400")
window.resizable(0, 0)

## Enabler Report Section
tkinter.Label(window, text = "Enabler Report Generator", font = ('arial', 12, 'bold')).grid(row =1, column = 2)
reportButton = tkinter.Button(window, text ="Generate Enabler Report", command = ES.startReportGeneration, height = 2, width = 25).grid(row = 2, column = 2)
mailButton  = tkinter.Button(window, text ="Send Enabler Report to Me", command = readMailcontent, height = 2, width = 25).grid(row = 3, column = 2)
teamMailButton  = tkinter.Button(window, text ="Send Enabler Report To Team", command = readMailcontentForTeam, height = 2, width = 25).grid(row = 4, column = 2)

## Health Report Section
tkinter.Label(window, text = "                  ", font = ('arial', 12, 'bold')).grid(row =0, column = 3)
tkinter.Label(window, text = "                  ", font = ('arial', 12, 'bold')).grid(row =0, column = 1)
tkinter.Label(window, text = "Health Report Generator", font = ('arial', 12, 'bold')).grid(row =1, column = 4)
healthReportButton = tkinter.Button(window, text ="Generate Health Report", command = BHR.generateBasicHealthReport, height = 2, width = 25).grid(row = 2, column = 4)
hMailButton  = tkinter.Button(window, text ="Send Health Report to Me", command =BRO.sendHealthReportToMeOnly , height = 2, width = 25).grid(row = 3, column = 4)
hTeamMailButton  = tkinter.Button(window, text ="Send Health Report To Team", command = '', height = 2, width = 25).grid(row = 4, column = 4)


tkinter.Label(window, text = " ", font = ('arial', 12, 'bold')).grid(row =6, column = 2)
tkinter.Label(window, text = "System Coverage Report", font = ('arial', 12, 'bold')).grid(row =7, column = 2)
coverageReportButton = tkinter.Button(window, text ="Generate Coverage Report", command = VRG.generateCoverageReport, height = 2, width = 35).grid(row = 8, column = 2)
coverageReportButton = tkinter.Button(window, text ="Generate Coverage Report (Doc Format)", command = lambda: VRG_DOC.generateVerificationReport(2), height = 2, width = 35).grid(row = 9, column = 2)


tkinter.Label(window, text = "System Verfication Report", font = ('arial', 12, 'bold')).grid(row =7, column = 4)
verificationReportButton = tkinter.Button(window, text ="Generate Verification Report (xlsFormat)", command = VRG.generateVerificationReport, height = 2, width = 35).grid(row = 8, column = 4)
verificationReportButtonDoc = tkinter.Button(window, text ="Generate Verification Report (Doc Format)", command = lambda: VRG_DOC.generateVerificationReport(1), height = 2, width = 35).grid(row = 9, column = 4)

#mailButton  = tkinter.Button(window, text ="Send Enabler Report to Me", command = readMailcontent, height = 2, width = 25).grid(row = 3, column = 2)
#teamMailButton  = tkinter.Button(window, text ="Send Enabler Report To Team", command = readMailcontentForTeam, height = 2, width = 25).grid(row = 4, column = 2)

window.mainloop()

