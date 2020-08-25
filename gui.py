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


## Health Report Section
tkinter.Label(window, text = "                  ", font = ('arial', 12, 'bold')).grid(row =0, column = 3)
tkinter.Label(window, text = "                  ", font = ('arial', 12, 'bold')).grid(row =0, column = 1)
tkinter.Label(window, text = "Health Report Generator", font = ('arial', 12, 'bold')).grid(row =1, column = 4)
healthReportButton = tkinter.Button(window, text ="Generate Health Report", command = BHR.generateBasicHealthReport, height = 2, width = 25).grid(row = 2, column = 4)
#hMailButton  = tkinter.Button(window, text ="Send Health Report to Me", command =BRO.sendHealthReportToMeOnly , height = 2, width = 25).grid(row = 3, column = 4)
#hTeamMailButton  = tkinter.Button(window, text ="Send Health Report To Team", command = '', height = 2, width = 25).grid(row = 4, column = 4)



window.mainloop()

