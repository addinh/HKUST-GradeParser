from tkinter import ttk
from tkinter import filedialog, messagebox

from pathlib import Path
import pandas as pd

from asgnApp import AsgnApp
from utility import Table, askcombobox

class HwApp(AsgnApp):
    def __init__(self, master = None):
        super().__init__(master)

        # tk.Frame stuff
        self.pack()
        self.padding = 10
        self.master.title('Homework Grade Parser')
        self.master.geometry('900x480')

        # DataFrames for ZINC scores and summary
        self.zinc: pd.DataFrame = None
        self.scores: pd.DataFrame = None

        # UI components
        self.canvasCSVLabel = ttk.Label(self, text='Select the Canvas CSV file:', width=40)
        self.canvasCSVLabel.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        self.canvasCSVButton = ttk.Button(self, text='Browse', command=self.canvasCSVButtonPressed)
        self.canvasCSVButton.grid(row=0, column=1, padx=10, pady=10)

        self.assignmentSelectionLabel = ttk.Label(self, text='Select assignment:')
        self.assignmentSelectionLabel.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

        self.assignmentSelectionCombobox = ttk.Combobox(self, state='readonly', width=30)
        self.assignmentSelectionCombobox.grid(row=1, column=1, padx=10, pady=10)
        self.assignmentSelectionCombobox.bind("<<ComboboxSelected>>", self.assignmentSelected)
        self.assignmentSelectionCombobox.config(state='disabled')

        self.gradeTable = Table(self)
        self.gradeTable.grid(row=2, column=0, rowspan=3, columnspan=3)

        self.zincButton = ttk.Button(self, text='Import grade sheet', command=self.zincButtonPressed)
        self.zincButton.grid(row=3, column=3, padx=10, pady=10)
        self.zincButton.config(state='disabled')

        self.generateButton = ttk.Button(self, text='Generate Canvas CSV', command=self.generateButtonPressed)
        self.generateButton.grid(row=5, column=1, padx=10, pady=10)
        self.generateButton.config(state='disabled')

    """
    Assignment selection event handler
    """
    def assignmentSelected(self, event):
        self.assignmentLabel = self.assignmentSelectionCombobox.get()
        self.zincButton.config(state='normal')
        self.updateTable()

    """
    Import ZINC report event handler
    """
    def zincButtonPressed(self):
        self.parseHWreport()
        self.zincButton.config(state='disabled')

    """
    Process homework gradefile.
    Accepts a single Excel file, which should contain one column with ITSC emails, and another column with the total score.
    There will be prompts for user to select these 2 columns if the default names are not found (SIS Login ID and Total).

    The total scores are parsed directly into the Canvas CSV for export, without being scaled.
    """
    def parseHWreport(self):
        # Open homework gradefile
        report_xlsx = filedialog.askopenfilename(filetypes=[('Excel files', '.xlsx .xls')], title='Select the homework .xlsx gradefile:')
        if report_xlsx == '':
            return
        
        # Default column names
        SID_COLUMN = 'SIS Login ID'
        TOT_COLUMN = 'Total'

        # If above columns are not found, ask for user specification
        self.report = pd.read_excel(report_xlsx, sheet_name=0)
        columns = list(self.report.columns)
        sid_col = SID_COLUMN
        tot_col = TOT_COLUMN
        if SID_COLUMN not in self.report.columns:
            sid_col = askcombobox('ITSC email column', 'Select the column containing ITSC emails:', columns)
        if TOT_COLUMN not in self.report.columns:
            tot_col = askcombobox('Total score column', 'Select the column containing total scores:', columns)
        self.report.rename(columns={sid_col: SID_COLUMN, tot_col:TOT_COLUMN}, inplace=True)

        # Fill scores into grades DataFrame
        self.report.dropna(subset=[SID_COLUMN], inplace=True)
        self.scores = self.report.loc[:,[SID_COLUMN, TOT_COLUMN]].rename(columns={TOT_COLUMN: self.assignmentLabel}).set_index(SID_COLUMN).reindex(self.grades[SID_COLUMN])
        self.grades.fillna(self.scores.reset_index(), inplace=True)
        self.grades[self.assignmentLabel].fillna(0, inplace=True)
        self.updateTable()

        # Enable output button
        self.generateButton.config(state='normal')
