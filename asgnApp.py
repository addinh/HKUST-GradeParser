import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox

from pathlib import Path
import pandas as pd

from utility import Table

class AsgnApp(tk.Frame):
    def __init__(self, master = None):
        super().__init__(master)

        # Path to Canvas CSV File
        self.grade_csv: str = None

        # Canvas CSV DataFrame
        self.grades: pd.DataFrame = None

        # Name of selected assignment in Canvas
        self.assignmentLabel: str = None

        # Calculated score DataFrame
        self.report: pd.DataFrame = None

        # List of additional columns to display on the table
        self.extraColumns: list[str] = []

        # Checks whether Canvas has Manual Posting enabled
        self.hasManualPostingRow: bool = True

        # To be initialized by derived classes
        self.canvasCSVLabel: ttk.Label = None
        self.assignmentSelectionCombobox: ttk.Combobox = None
        self.gradeTable: Table = None


    @property
    def tableColumns(self) -> list:
        return ['Student', 'SIS User ID', 'SIS Login ID', 'Section', self.assignmentLabel]

    @property
    def assignmentName(self) -> str:
        if not self.assignmentLabel: 
            return None
        return self.assignmentLabel.split(' (')[0]

    @property
    def table(self) -> pd.DataFrame:
        if self.grades is None or not self.assignmentLabel:
            return None
        if self.report is None:
            return self.grades.loc[:, self.tableColumns]
        table = self.grades.merge(self.report.loc[:, ['SIS Login ID'] + self.extraColumns], how='left', on='SIS Login ID')
        return table.loc[:, self.tableColumns + self.extraColumns]
    
    """
    Import Canvas CSV event handler
    """
    def canvasCSVButtonPressed(self):
        grade_csv = filedialog.askopenfilename(filetypes=[('CSV files', '.csv')], title='Select the Canvas Grade Export .csv file:')
        if not grade_csv:
            return
        self.grade_csv = grade_csv
        self.canvasCSVLabel.config(text='Select the Canvas CSV file:\n{}'.format(Path(self.grade_csv).name))

        self.grades = pd.read_csv(self.grade_csv)
        self.grades = self.grades.astype({'ID': 'Int64', 'SIS User ID': 'Int64'})

        self.hasManualPostingRow = (self.grades['Student'][0] != '    Points Possible')
        if not self.hasManualPostingRow:
            messagebox.showwarning(title='Warning', message='Grade Posting Policy detected as Automatic. Consider changing it on Canvas.')

        gradesNanCount = self.grades.drop(index=0).isna().sum()
        possibleAssignments = gradesNanCount.loc[gradesNanCount == gradesNanCount.max()].index
        self.assignmentSelectionCombobox.config(values=list(possibleAssignments), state='normal')

    """
    Export Canvas CSV event handler
    """
    def generateButtonPressed(self):
        # Output to CSV
        output_csv = Path(self.grade_csv).parent / '{}_parsedGrade.csv'.format(self.assignmentName)
        self.grades.to_csv(output_csv, index=False, float_format='%.2f')
        messagebox.showinfo(title='Finished processing', message='Written to "{}". Import this file to Canvas Gradebook.'.format(output_csv))

    """
    Helper function to update content of table
    """
    def updateTable(self):
        self.gradeTable.initTable(self.table.shape[0] + 1 - (3 if self.hasManualPostingRow else 2), self.table.shape[1])
        self.gradeTable.setDataframe(self.table, 2 if self.hasManualPostingRow else 1)