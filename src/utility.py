import tkinter as tk
from tkinter import ttk, messagebox, simpledialog

import pandas as pd
from math import isnan

class ScrollbarFrame(tk.Frame):
    """
    Extends class tk.Frame to support a scrollable Frame 
    This class is independent from the widgets to be scrolled and 
    can be used to replace a standard tk.Frame
    """
    def __init__(self, parent, **kwargs):
        tk.Frame.__init__(self, parent, **kwargs)

        # The Scrollbar, layout to the right
        vsb = tk.Scrollbar(self, orient="vertical")
        vsb.pack(side="right", fill="y")

        # The Scrollbar, layout to the right
        hsb = tk.Scrollbar(self, orient="horizontal")
        hsb.pack(side="bottom", fill="x")

        # The Canvas which supports the Scrollbar Interface, layout to the left
        self.canvas = tk.Canvas(self, borderwidth=0, width=600, height=300)
        self.canvas.pack(side="left", fill="both", expand=True)

        # Bind the Scrollbar to the self.canvas Scrollbar Interface
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.configure(command=self.canvas.yview)
        hsb.configure(command=self.canvas.xview)

        # The Frame to be scrolled, layout into the canvas
        # All widgets to be scrolled have to use this Frame as parent
        self.scrolled_frame = tk.Frame(self.canvas, background='black')
        self.canvas.create_window((4, 4), window=self.scrolled_frame, anchor="nw")

        # Configures the scrollregion of the Canvas dynamically
        self.scrolled_frame.bind("<Configure>", self.on_configure)

    def on_configure(self, event):
        """Set the scroll region to encompass the scrolled frame"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


class Table(ScrollbarFrame):
    """
    Displays a table on the GUI.
    Supports scrolling horizontally and vertically.
    """
    def __init__(self, parent, rows=0, columns=0):
        ScrollbarFrame.__init__(self, parent)
        self._widgets: list[list[tk.Label]] = []
        self.rows = rows
        self.columns = columns
        self.initTable(rows, columns)

    def get(self, row, column):
        widget = self._widgets[row][column]
        return widget.cget('text')

    def set(self, row, column, value):
        widget = self._widgets[row][column]
        if (isinstance(value, float) and not isnan(value)):
            widget.config(text='{:.2f}'.format(value))
        else:
            widget.config(text=value)

    def setColumn(self, column, values):
        row = 1
        for value in values:
            self.set(row, column, value)
            row += 1
    
    def setDataframe(self, df: pd.DataFrame, startIndex):
        columnNames = df.columns.tolist()
        for i in range(df.shape[1]):
            self.set(0, i, columnNames[i])
            self.setColumn(i, df[columnNames[i]][startIndex:-1])
    
    def initTable(self, rows, columns):
        for row in range(self.rows):
            for column in range(self.columns):
                widget = self._widgets[row][column]
                widget.destroy()
        self._widgets = []
        self.rows = rows
        self.columns = columns
        for row in range(rows):
            current_row = []
            for column in range(columns):
                label = tk.Label(self.scrolled_frame, text="%s/%s" % (row, column), 
                                 borderwidth=1)
                label.grid(row=row, column=column, sticky='nsew', padx=1, pady=1)
                current_row.append(label)
            self._widgets.append(current_row)

        for column in range(columns):
            self.grid_columnconfigure(column, weight=1)


class ComboboxDialog(simpledialog.Dialog):
    """
    Extends Dialog class to support a Dialog with Combobox prompt.
    """
    def __init__(self, title, prompt,
                 values,
                 parent = None):

        self.prompt   = prompt
        self.values   = values
        simpledialog.Dialog.__init__(self, parent, title)

    def destroy(self):
        self.entry = None
        simpledialog.Dialog.destroy(self)

    def body(self, master):

        w = ttk.Label(master, text=self.prompt, justify='left')
        w.grid(row=0, padx=5, sticky='w')

        self.entry = ttk.Combobox(master, name="entry", values=self.values)
        self.entry.grid(row=1, padx=5, sticky='we')

        return self.entry

    def validate(self):
        try:
            result = self.getresult()
        except ValueError:
            messagebox.showwarning(
                "Illegal value\nPlease try again",
                parent = self
            )
            return 0

        self.result = result

        return 1
    
    def getresult(self):
        return self.entry.get()
    
"""
Helper function to create a Combobox and return the result
"""
def askcombobox(title, prompt, values, **kw):
    d = ComboboxDialog(title, prompt, values, **kw)
    return d.result