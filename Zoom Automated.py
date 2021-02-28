import pyautogui
import datetime
import pandas
import time
import tkinter
from tkinter import *

def Execute():
    print("Initiating Process Zoom")
    time.sleep(5)
    #Opening Zoom from deskstop
    pyautogui.moveTo(22,23)
    pyautogui.doubleClick()
    print("Step 1 complete, moving to step 2")
    
reader = pandas.read_excel("E:\Zoom Automated\Zoom Credentials.xls")

def English():
    meetingid1 = str(reader.iloc[0,1])
    #Accessing meeting ID from row 2
    passcode1 = str(reader.iloc[0,2])
    print("Step 2 complete, moving to step 3")
    time.sleep(6)
    #Pressing Join Button
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.write(meetingid1)
    #Parsing data to Zoom meetingid field
    time.sleep(5)
    pyautogui.press("enter")
    print("Step 3 complete, moving to step 4")
    time.sleep(7)
    #Parsing Data to Zoom Password field
    pyautogui.write(passcode1)
    time.sleep(2)
    pyautogui.press("enter")
    #Joining the meeting
    time.sleep(10)
    #Maximize Zoom Window
    pyautogui.moveTo(22,23)
    pyautogui.doubleClick()
    print("Step 4 of 4 Successfully Completed.")

def Biology():
    meetingid2 = str(reader.iloc[1,1])
    #Accessing meeting ID from row 2
    passcode2 = str(reader.iloc[1,2])
    print("Step 2 complete, moving to step 3")
    time.sleep(6)
    #Pressing Join Button
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.write(meetingid2)
    #Parsing data to Zoom meetingid field
    time.sleep(5)
    pyautogui.press("enter")
    print("Step 3 complete, moving to step 4")
    time.sleep(7)
    #Parsing Data to Zoom Password field
    pyautogui.write(passcode2)
    time.sleep(2)
    pyautogui.press("enter")
    #Joining the meeting
    time.sleep(10)
    #Maximize Zoom Window
    pyautogui.moveTo(22,23)
    pyautogui.doubleClick()
    print("Step 4 of 4 Successfully Completed.")

def Maths():
    meetingid3 = str(reader.iloc[2,1])
    #Accessing meeting ID from row 2
    passcode3 = str(reader.iloc[2,2])
    print("Step 2 complete, moving to step 3")
    time.sleep(6)
    #Pressing Join Button
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.write(meetingid3)
    #Parsing data to Zoom meetingid field
    time.sleep(5)
    pyautogui.press("enter")
    print("Step 3 complete, moving to step 4")
    time.sleep(7)
    #Parsing Data to Zoom Password field
    pyautogui.write(passcode3)
    time.sleep(2)
    pyautogui.press("enter")
    #Joining the meeting
    time.sleep(10)
    #Maximize Zoom Window
    pyautogui.moveTo(22,23)
    pyautogui.doubleClick()
    print("Step 4 of 4 Successfully Completed.")

def Chemistry():
    meetingid3 = str(reader.iloc[3,1])
    #Accessing meeting ID from row 2
    passcode3 = str(reader.iloc[3,2])
    print("Step 2 complete, moving to step 3")
    time.sleep(6)
    #Pressing Join Button
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.write(meetingid3)
    #Parsing data to Zoom meetingid field
    time.sleep(5)
    pyautogui.press("enter")
    print("Step 3 complete, moving to step 4")
    time.sleep(7)
    #Parsing Data to Zoom Password field
    pyautogui.write(passcode3)
    time.sleep(2)
    pyautogui.press("enter")
    #Joining the meeting
    time.sleep(10)
    #Maximize Zoom Window
    pyautogui.moveTo(22,23)
    pyautogui.doubleClick()
    print("Step 4 of 4 Successfully Completed.")

def Physics():
    meetingid4 = str(reader.iloc[4,1])
    #Accessing meeting ID from row 2
    passcode4 = str(reader.iloc[4,2])
    print("Step 2 complete, moving to step 3")
    time.sleep(6)
    #Pressing Join Button
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.write(meetingid4)
    #Parsing data to Zoom meetingid field
    time.sleep(5)
    pyautogui.press("enter")
    print("Step 3 complete, moving to step 4")
    time.sleep(7)
    #Parsing Data to Zoom Password field
    pyautogui.write(passcode4)
    time.sleep(2)
    pyautogui.press("enter")
    #Joining the meeting
    time.sleep(10)
    #Maximize Zoom Window
    pyautogui.moveTo(22,23)
    pyautogui.doubleClick()
    print("Step 4 of 4 Successfully Completed.")

def ExecuteEnglish():
    Execute()
    English()

def ExecuteBiology():
    Execute()
    Biology()

def ExecuteMaths():
    Execute()
    Maths()

def ExecuteChemistry():
    Execute()
    Chemistry()

def ExecutePhysics():
    Execute()
    Physics()


prog = Tk()
prog.geometry("350x350")
prog.title("Zoom Automator")
Label( prog, text = " Zoom Bot", font = " sansserif 16 italic").pack()
Label( prog, text = "Select Class", font = " sansserif 10 italic").pack()
Label( prog, text = " Proudly made by Anas Alwi ", font = " sansserif 10 italic").pack( side = "bottom")

btn1 = Button(prog, text ="English", bg = "White", activebackground = "Yellow", activeforeground = "White", font = "sansserif 8 italic", bd = 1, command = ExecuteEnglish)
btn2 = Button(prog, text ="Biology", bg = "White", activebackground = "Yellow", activeforeground = "White", font = "sansserif 8 italic", bd = 1, command= ExecuteBiology)
btn3 = Button(prog, text ="Maths", bg = "White", activebackground = "Yellow", activeforeground = "White", font = "sansserif 8 italic", bd = 1, command = ExecuteMaths)
btn4 = Button(prog, text ="Chemistry", bg = "White", activebackground = "Yellow", activeforeground = "White", font = "sansserif 8 italic", bd = 1, command = ExecuteChemistry)
btn5 = Button(prog, text ="Physics", bg = "White", activebackground = "Yellow", activeforeground = "White", font = "sansserif 8 italic", bd = 1, command = ExecutePhysics)

btn1.pack(side= "top")
btn2.pack(side= "top")
btn3.pack(side= "top")
btn4.pack(side= "top")
btn5.pack(side= "top")

prog.mainloop()
