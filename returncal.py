import os
from tkinter import *
import tkinter.filedialog as fl
import tax
import helper
root = Tk()
root.title('Tax Returns')
root.geometry('300x300')
wblist =[]


wbno = -1
def disptext(text):
    filewin = Toplevel(root)
    label = Label(filewin, text=text, padx =5 ,pady=4)
    button = Button(filewin, text='ok',command =filewin.destroy,padx =5 ,pady=1)
    label.pack()
    button.pack()

def createbuttons():
    global button1
    global button2
    button1= Button(root, text = 'remove characters', command = sheetmod, padx =2, pady =2)
    button2= Button(root, text = 'cacluate value', command = recal, padx =2, pady =2)
    button1.pack()
    button2.pack()
    
def sheetopen():
    location = fl.askopenfilename()
    wblist.append(tax.openwb(location))
    global wbno
    wbno+=1
    createbuttons()

    
def sheetmod():
    tax.replace(wblist[wbno],6,2)
    disptext('special characters removed')


def recal():
    tax.calculateret(wblist[wbno],6,5)
    disptext('calculation done')
    

def save():
    tax.save(wblist[wbno],'output.xlsx')
    disptext('spreadsheet saved')

def close():
    save()
    button1.destroy()
    button2.destroy()
    disptext('open new spreadsheet')

def abou():
    text = helper.about()
    disptext(text)
def hel():
    text = helper.helper()
    disptext(text)    
    
def exit1():
    root.destroy()
    
menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Open", command=sheetopen)
filemenu.add_command(label="Save", command=save)
filemenu.add_command(label="Close", command=close)

filemenu.add_separator()

filemenu.add_command(label="Exit", command=exit1)
menubar.add_cascade(label="File", menu=filemenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Help Index" , command = hel )
helpmenu.add_command(label="About..." , command  = abou )
menubar.add_cascade(label="Help", menu=helpmenu)

root.config(menu=menubar)



root.mainloop()
