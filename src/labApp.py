from tkinter import ttk
from tkinter import filedialog

import pandas as pd

from asgnApp import AsgnApp
from utility import Table, askcombobox

class LabApp(AsgnApp):
    def __init__(self, master = None):
        super().__init__(master)

        # tk.Frame stuff
        self.pack()
        self.padding = 10
        self.master.title('COMP2012 Lab Grade Parser')
        self.master.geometry('900x480')

        # DataFrames for Attendance, ZINC scores, Question scores and summary
        self.attendance: pd.DataFrame = None
        self.zinc: pd.DataFrame = None
        self.question: pd.DataFrame = None
        self.scores: pd.DataFrame = None

        # Flags to check all 3 components have been imported
        self.attendanceFlag: bool = False
        self.zincFlag: bool = False
        self.questionFlag: bool = False

        # Modify extraColumns for COMP2012 Lab
        self.extraColumns = ['Attendance', 'Lucky?', 'Question score', 'ZINC']

        # UI components
        self.canvasCSVLabel = ttk.Label(self, text='Select the Canvas CSV file:', width=40)
        self.canvasCSVLabel.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        self.canvasCSVButton = ttk.Button(self, text='Browse', command=self.canvasCSVButtonPressed)
        self.canvasCSVButton.grid(row=0, column=1, padx=10, pady=10)

        self.numLabSessionLabel = ttk.Label(self, text='Number of lab sessions:')
        self.numLabSessionLabel.grid(row=0, column=2, padx=10, pady=10, sticky='ew')

        self.numLabSessionSpinbox = ttk.Spinbox(self, from_=1, to=20)
        self.numLabSessionSpinbox.grid(row=0, column=3, padx=10, pady=10)
        self.numLabSessionSpinbox.set(1)

        self.zincMaxLabel = ttk.Label(self, text='Maximum ZINC Score:')
        self.zincMaxLabel.grid(row=1, column=2, padx=10, pady=10, sticky='ew')

        self.zincMaxEntry = ttk.Entry(self)
        self.zincMaxEntry.grid(row=1, column=3, padx=10, pady=10)
        self.zincMaxEntry.insert(0, '100')

        self.assignmentSelectionLabel = ttk.Label(self, text='Select assignment:')
        self.assignmentSelectionLabel.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

        self.assignmentSelectionCombobox = ttk.Combobox(self, state='readonly', width=30)
        self.assignmentSelectionCombobox.grid(row=1, column=1, padx=10, pady=10)
        self.assignmentSelectionCombobox.bind("<<ComboboxSelected>>", self.assignmentSelected)
        self.assignmentSelectionCombobox.config(state='disabled')

        self.gradeTable = Table(self)
        self.gradeTable.grid(row=2, column=0, rowspan=3, columnspan=3)

        self.attendanceButton = ttk.Button(self, text='Import Attendance', command=self.attendanceButtonPressed)
        self.attendanceButton.grid(row=2, column=3, padx=10, pady=10)
        self.attendanceButton.config(state='disabled')

        self.zincButton = ttk.Button(self, text='Import ZINC report(s)', command=self.zincButtonPressed)
        self.zincButton.grid(row=3, column=3, padx=10, pady=10)
        self.zincButton.config(state='disabled')

        self.questionButton = ttk.Button(self, text='Import lab question score', command=self.questionButtonPressed)
        self.questionButton.grid(row=4, column=3, padx=10, pady=10)
        self.questionButton.config(state='disabled')

        self.generateButton = ttk.Button(self, text='Generate Canvas CSV', command=self.generateButtonPressed)
        self.generateButton.grid(row=5, column=1, padx=10, pady=10)
        self.generateButton.config(state='disabled')
    
    @property
    def zincMax(self) -> float:
        try:
            return float(self.zincMaxEntry.get())
        except ValueError:
            print('ZINC max score cannot be parsed. Using default value of 100.')
            return 100

    """
    Assignment selection event handler
    """
    def assignmentSelected(self, event):
        self.assignmentLabel = self.assignmentSelectionCombobox.get()
        self.attendanceButton.config(state='normal')
        self.questionButton.config(state='normal')
        self.zincButton.config(state='normal')
        self.updateTable()

    """
    Import attendance event handler
    """
    def attendanceButtonPressed(self):
        self.parseLabAttendance()
        self.attendanceFlag = True
        if self.attendanceFlag and self.zincFlag and self.questionFlag:
            self.processLabScores()
        self.attendanceButton.config(state='disabled')

    """
    Import ZINC reports event handler
    """
    def zincButtonPressed(self):
        self.parseLabZINCreports()
        self.zincFlag = True
        if self.attendanceFlag and self.zincFlag and self.questionFlag:
            self.processLabScores()
        self.zincButton.config(state='disabled')

    """
    Import question score event handler
    """
    def questionButtonPressed(self):
        self.parseLabQuestions()
        self.questionFlag = True
        if self.attendanceFlag and self.zincFlag and self.questionFlag:
            self.processLabScores()
        self.questionButton.config(state='disabled')

    """
    Process attendance sheet.
    Supports multiple Excel files, though only one should be needed under current arrangement.
    The Excel file(s) should contain 1 sheet named 'Tally', containing Email, Name and Score columns. 
    """
    def parseLabAttendance(self):
        # Open attendance reports
        attendance_xlsxs = filedialog.askopenfilenames(filetypes=[('Excel files', '.xlsx .xls')], title='Select the attendance .xlsx gradefiles:')
        if attendance_xlsxs == '':
            return

        # Parse all reports into a single DataFrame
        attendances = [pd.read_excel(attendance_xlsx, sheet_name='Tally') for attendance_xlsx in attendance_xlsxs]
        self.attendance = pd.concat(attendances)
        self.attendance.drop_duplicates(subset=['Email'], inplace=True)

    """
    Process ZINC reports.
    Supports multiple Excel files, downloaded from ZINC. Each lab session has one submission module for different due dates.
    The Excel files should contain 1 sheet with ITSC, Name and Score columns.

    If there are duplicate ITSCs when combining sheets, a message box will ask for which score to keep. 
    This happens if the TA submits to all sessions for checking (in which case select any option),
    or a student was able to submit to multiple sessions due to lab swap. Manually check ZINC to see which one is more recent.

    Automatically generates Email column for compatibility, and scale the ZINC score based on user-specified maximum score.
    It is recommended to always set ZINC lab score to 100 maximum, so that you don't accidentally forget this step. 
    """
    def parseLabZINCreports(self):
        # Open ZINC reports
        zinc_xlsxs = filedialog.askopenfilenames(filetypes=[('Excel files', '.xlsx .xls')], title='Select the ZINC .xlsx gradefiles:')
        if zinc_xlsxs == '':
            return

        # Parse all reports into a single DataFrame
        zincs = [pd.read_excel(zinc_xlsx, sheet_name=0) for zinc_xlsx in zinc_xlsxs]
        self.zinc = pd.concat(zincs)

        # Handle duplicates
        self.zinc.reset_index(inplace=True)
        self.zinc['Summary'] = self.zinc.apply(lambda row: '{}: Score: {}, Late: {}'.format(row['Name'], row['Score'], row['Late Submission']), axis=1)
        duplicates = self.zinc.loc[self.zinc.duplicated(subset=['ITSC'], keep=False)]
        itscs = duplicates['ITSC'].drop_duplicates()
        for itsc in itscs:
            options = []
            for _, row in duplicates[duplicates['ITSC'] == itsc].iterrows():
                options.append(row['Summary'])
            keep = askcombobox('Duplicate', 'Select the score you want to keep for student {}'.format(itsc), options)
            index = self.zinc[(self.zinc['ITSC'] == itsc) & (self.zinc['Summary'] != keep)].index
            self.zinc.drop(index=index, inplace=True)

        # Cleanup
        self.zinc.drop_duplicates(subset=['ITSC'], inplace=True)
        self.zinc['Email'] = self.zinc['ITSC'] + '@connect.ust.hk'
        self.zinc['ZINC'] = self.zinc['Score'].div(self.zincMax)

        # Change default number of labs
        self.numLabSessionSpinbox.set(len(zincs))
    
    """
    Process question score sheet.
    Accepts a single Excel file with the first N sheets containing question score for each lab session, N for number of lab sessions.
    The number of lab sessions is automatically set if ZINC reports are imported first.
    Each sheet should contain Email, Name, 'Lucky?' and 'Question score' columns.
    """
    def parseLabQuestions(self):
        # Open question report
        question_xlsx = filedialog.askopenfilename(filetypes=[('Excel files', '.xlsx .xls')], title='Select the question .xlsx gradefile:')
        numLabs = int(self.numLabSessionSpinbox.get())
        if question_xlsx == '':
            return

        # Parse all tabs into a single DataFrame
        questions = [pd.read_excel(question_xlsx, sheet_name=i) for i in range(numLabs)]
        self.question = pd.concat(questions)
        self.question.drop_duplicates(subset=['Email'], inplace=True)

    """
    Calculate lab scores when all 3 components have been imported.
    Formula: Attendance + ZINC + (Question if Lucky else ZINC)

    Scores are then imported into the self.grades DataFrame, with absent students receiving 0.
    The scores will be viewable on the table and the output CSV is ready to be exported.
    """
    def processLabScores(self):
        # Merge and process score
        self.report = self.attendance.merge(self.question.loc[:, ['Email', 'Lucky?', 'Question score']], how='left', on='Email')
        self.report = self.report.merge(self.zinc.loc[:, ['Email', 'ZINC']], how='left', on='Email')
        self.report['Question score'].fillna(0, inplace=True)
        self.report['ZINC'].fillna(0, inplace=True)
        self.report['Total'] = self.report.apply(lambda row: row['Attendance'] + row['ZINC'] + (row['Question score'] if row['Lucky?'] == 'Yes' else row['ZINC']), axis=1)
        self.report['Total'] = self.report['Total'].round(2).apply(str)
        self.report.rename(columns={'Email': 'SIS Login ID', 'Total': self.assignmentLabel}, inplace=True)

        # Fill scores into grades DataFrame
        self.scores = self.report.loc[:,['SIS Login ID', self.assignmentLabel]].set_index('SIS Login ID').reindex(self.grades['SIS Login ID'])
        self.grades.fillna(self.scores.reset_index(), inplace=True)
        self.grades[self.assignmentLabel].fillna(0, inplace=True)
        self.updateTable()

        # Enable output button
        self.generateButton.config(state='normal')
