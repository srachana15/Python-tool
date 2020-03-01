# -*- coding: utf-8 -*-
"""
Spyder Editor

"""

import os
import re
import zipfile
import tkinter
from tkinter import filedialog, messagebox

def show_error(text):
    tkinter.messagebox.showerror('Error',text)

def show_results(text):
    tkinter.messagebox.showinfo('Results',text)
    
def count(files):
    #define a dictionary to hold the filename and a count of number of slides:
    decks = {}
    results = ""
    #for each file check if it is a pptx file and set count to 0 or show error message
    for file in files:
        if os.path.abspath(file).endswith('.pptx'):
            decks[(os.path.abspath(file))] = 0
        else:
            show_error('The file %s is not a .pptx file and it will be ignored'%(file))
    #counting slides in each deck
    for deck, count in decks.items():
        try:
            #try to read the file as a zipfile since pptx works like a zipfile
            archive = zipfile.ZipFile(deck,'r')
            contents = archive.namelist()
            #We cannot directly count the len of contents to get the number of slides as the list contains elemetns like 'ppt','slides','docprops',etc.
        except Exception as e:
            #error in reading file
            show_error('Error reading %s (%s). Hence, count set to 0'%(os.path.abspath(deck),e))
        #if no exeption occurred 
        else:
            for name in contents:
                if(re.findall('ppt/slides/slide',name)):
                    decks[deck] += 1
    results += ('Slides\tDeck\n')
    
    for deck, count in decks.items():
        results += ('%s\t\t%s\n'%(count,os.path.basename(deck)))
    
    if show_summary.get():
        results += ('\n- - - - -\n')
        total = 0
        for count in decks.values():
            total += count
        results += ('Total slides in %s decks: %s'%(len(decks),total))
    show_results(results)

#Creating the GUI to read the pptx files
root = tkinter.Tk()
root.geometry('275x175')
root.title("Slide Counter")

#create file dialogbox to select files
def openfilebutton():
    t = tkinter.filedialog.askopenfilenames()
    count(t)

#Design the dialog box
header = tkinter.Label(root,text='Welcome to Slide Counter!', fg='blue', font=('Times Bold', 18))
header.pack(side='top',ipady = 10)

text = tkinter.Label(root,text='Select some .pptx files, and the\n app will count how many slides\n are contained within them.')
text.pack()

show_summary = tkinter.BooleanVar() #create a variable
show_summary.set(True)
summary = tkinter.Checkbutton(root, text='Show summary', var=show_summary) #create a checkbutton and link it to the bool var
summary.pack(ipady = 10)

openfiles = tkinter.Button(root, text='Choose files', command=openfilebutton)
openfiles.pack(fill="x")

# Initialize Tk window.
root.mainloop()
