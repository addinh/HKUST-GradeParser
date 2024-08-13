from tkinter import ttk
from tkinter import filedialog, messagebox

from pathlib import Path
import os
import pandas as pd
from matplotlib import pyplot as plt
from math import isnan

from asgnApp import AsgnApp
from utility import Table, askcombobox

class PaApp(AsgnApp):
    def __init__(self, master = None):
        super().__init__(master)

        # tk.Frame stuff
        self.pack()
        self.padding = 10
        self.master.title('COMP2012 PA Grade Parser')
        self.master.geometry('900x480')

        # DataFrames for ZINC scores and summary
        self.zinc: pd.DataFrame = None
        self.scores: pd.DataFrame = None

        # Modify extraColumns for COMP2012 PA
        self.extraColumns = ['Score', 'Penalty']

        # UI components
        self.canvasCSVLabel = ttk.Label(self, text='Select the Canvas CSV file:', width=40)
        self.canvasCSVLabel.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        self.canvasCSVButton = ttk.Button(self, text='Browse', command=self.canvasCSVButtonPressed)
        self.canvasCSVButton.grid(row=0, column=1, padx=10, pady=10)

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

        self.zincButton = ttk.Button(self, text='Import ZINC report(s)', command=self.zincButtonPressed)
        self.zincButton.grid(row=3, column=3, padx=10, pady=10)
        self.zincButton.config(state='disabled')

        self.statsButton = ttk.Button(self, text='Generate Stats', command=self.statsButtonPressed)
        self.statsButton.grid(row=5, column=0, padx=10, pady=10)
        self.statsButton.config(state='disabled')

        self.generateButton = ttk.Button(self, text='Generate Canvas CSV', command=self.generateButtonPressed)
        self.generateButton.grid(row=5, column=1, padx=10, pady=10)
        self.generateButton.config(state='disabled')

        self.jplagButton = ttk.Button(self, text='Generate JPlag report', command=self.jplagButtonPressed)
        self.jplagButton.grid(row=5, column=2, padx=10, pady=10)
        self.jplagButton.config(state='disabled')
    
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
        self.zincButton.config(state='normal')
        self.updateTable()

    """
    Import ZINC report event handler
    """
    def zincButtonPressed(self):
        self.parsePAreport()
        self.zincButton.config(state='disabled')

    """
    Process ZINC report.
    Supports multiple Excel files, downloaded from ZINC.
    Typically only 1 file is needed, but in some cases there are multiple ZINC submission modules with different deadlines.

    If there are duplicate ITSCs when combining sheets, a message box will ask for which score to keep.
    Check ZINC to see which one is more recent if needed.

    Automatically calculates total PA score using the formula: Total = max(ZINC / zincMax - Penalty, 0) where:
    - Penalty is the number of late minutes, rounded down
    - zincMax is set by the user *before* importing the ZINC reports

    Scores are then imported into the self.grades DataFrame, with absent students receiving 0.
    The scores will be viewable on the table and the output CSV is ready to be exported.
    """
    def parsePAreport(self):
        # Open reports
        zinc_xlsxs = filedialog.askopenfilenames(filetypes=[('Excel files', '.xlsx .xls')], title='Select the ZINC .xlsx gradefiles:')
        if zinc_xlsxs == '':
            return

        # Parse all reports into a single DataFrame
        zincs = [pd.read_excel(zinc_xlsx, sheet_name=0) for zinc_xlsx in zinc_xlsxs]
        self.zinc = pd.concat(zincs)
        
        # Handle duplicates
        self.zinc.reset_index(drop=True, inplace=True)
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
        self.zinc.drop_duplicates(subset=['ITSC'], inplace=True)
        self.zinc.drop(columns=['Summary'], inplace=True)

        # Process report
        self.report = self.zinc.loc[:,['ITSC', 'Name', 'Score', 'Late Submission']]
        self.report['Penalty'] = pd.to_numeric(self.report['Late Submission'].fillna('0.').str.split('.').str.get(0))
        self.report['Total'] = (self.report['Score'] / self.zincMax * 100 - self.report['Penalty']).clip(lower=0)
        self.report['Total'] = self.report['Total'].round(2).apply(str)
        self.report['SIS Login ID'] = self.report['ITSC'] + '@connect.ust.hk'
        self.report.rename(columns={'Total': self.assignmentLabel}, inplace=True)

        # Fill scores into grades DataFrame
        self.scores = self.report.loc[:,['SIS Login ID', self.assignmentLabel]].set_index('SIS Login ID').reindex(self.grades['SIS Login ID'])
        self.scores = self.scores.astype({self.assignmentLabel: 'Float64'})
        self.grades.fillna(self.scores.reset_index(), inplace=True)
        self.grades[self.assignmentLabel].fillna(0, inplace=True)
        self.updateTable()

        # Enable output button(s)
        self.statsButton.config(state='normal')
        self.generateButton.config(state='normal')
        self.jplagButton.config(state='normal')
    
    """
    Generate stats for PA.
    - Score histogram
    - Score distribution & optional distribution by test case
    - Lookup .csv file
    """
    def statsButtonPressed(self):
        # Ask for output directory
        output_dir = filedialog.askdirectory(title='Select output directory for the stats:')
        if output_dir == '':
            output_dir = Path(self.grade_csv).parent

        # Generate histogram
        output_hist = Path(output_dir) / 'histogram.png'
        fig = plt.hist(self.grades[self.assignmentLabel][(2 if self.hasManualPostingRow else 1):-1].astype('Float64'), bins=range(0, 101, 10), rwidth=0.8)
        plt.title(self.assignmentName + ' Histogram')
        plt.ylabel('Number of students')
        plt.xlabel('Score')
        plt.savefig(output_hist)
        plt.close()

        # Generate stats
        stats_txt = Path(output_dir) / '{}_stats.txt'.format(self.assignmentName)
        testcases = self.zinc.columns[4:]
        with open(stats_txt, mode='w') as sf:
            print('Mean: {:.2f}'.format(self.scores[self.assignmentLabel].mean()), file=sf)
            print('SD: {:.2f}'.format(self.scores[self.assignmentLabel].std()), file=sf)
            print('Max: {:.2f}'.format(self.scores[self.assignmentLabel].max()), file=sf)
            print('Median: {:.2f}'.format(self.scores[self.assignmentLabel].median()), file=sf)
            print('Min: {:.2f}'.format(self.scores[self.assignmentLabel].min()), file=sf)
            print(file=sf)
            testcaseStats = self.zinc[testcases].astype('Float64').sum() \
                                .div(self.zinc.shape[0]).div(self.zinc[testcases].astype('Float64').max()) \
                                .to_frame().rename(columns={0: 'Pass percentage'})
            testcaseStats['Visualization'] = testcaseStats.apply(lambda row: '[{}]'.format(''.join(['=' if i/40 < row['Pass percentage'] else ' ' for i in range(40)])), axis=1)
            print(testcaseStats.to_string(), file=sf)

        # Generate lookup CSV
        lookup_csv = Path(output_dir) / 'lookup' / 'score.csv'
        if not os.path.exists(lookup_csv.parent):
            os.mkdir(lookup_csv.parent)
        lookup = self.table.loc[:, ['Student', 'SIS User ID', 'SIS Login ID', 'Score', 'Penalty', self.assignmentLabel]]
        lookup.rename(columns={self.assignmentLabel: 'Total'}, inplace=True)
        lookup.dropna(subset=['SIS User ID'], inplace=True)
        lookup['Remarks'] = lookup.apply(lambda row: 'No submission' if isnan(row['Score']) else '', axis=1)
        lookup.to_csv(lookup_csv, index=False)

        messagebox.showinfo(title='Finished processing', message='Stats written to "{}".'.format(output_dir))

    """
    Generate JPlag report
    TODO
    """
    def jplagButtonPressed(self):
        messagebox.showwarning(message='This option will be implemented later!')
