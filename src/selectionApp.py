import tkinter as tk
from tkinter import ttk

from labApp import LabApp
from paApp import PaApp
from hwApp import HwApp

COMP2012_LAB = 1
COMP2012_PA  = 2
GENERAL_HW  = 3
ASSIGNMENT_DICT = {
    'COMP2012 Lab': COMP2012_LAB, 
    'COMP2012 PA': COMP2012_PA, 
    'General Homework': GENERAL_HW
}

class selectionApp(tk.Frame):
    def __init__(self, master = None):
        super().__init__(master)
        self.pack()
        self.padding = 10

        self.master.title('COMP2012/2611 Grade Parser')
        self.master.geometry('600x200')

        self.assignmentTypeLabel = ttk.Label(self, text='Select the type of assignment:')
        self.assignmentTypeLabel.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        self.assignmentTypeCombobox = ttk.Combobox(self, values=list(ASSIGNMENT_DICT.keys()), state='readonly', width=30)
        self.assignmentTypeCombobox.grid(row=1, column=0, padx=10, pady=10)
        self.assignmentTypeCombobox.bind("<<ComboboxSelected>>", self.assignmentTypeSelected)

    def assignmentTypeSelected(self, event):
        assignmentType = ASSIGNMENT_DICT[self.assignmentTypeCombobox.get()]

        self.master.destroy()
        asgnApp = None

        if (assignmentType == COMP2012_LAB):
            asgnApp = LabApp()
        elif (assignmentType == COMP2012_PA):
            asgnApp = PaApp()
        elif (assignmentType == GENERAL_HW):
            asgnApp = HwApp()

        if asgnApp is not None:
            asgnApp.mainloop()

