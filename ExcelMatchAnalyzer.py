import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Scrollbar, Listbox, VERTICAL, END, Checkbutton, IntVar
import os

class ExcelComparator:
    """
    A class to handle the comparison of two Excel files based on user-selected columns.
    """

    def __init__(self, root):
        """
        Initialize the ExcelComparator with the main Tkinter window and create the GUI.

        :param root: The main Tkinter window.
        """
        self.root = root
        self.root.title('Excel File Comparator')
        self.file1 = None
        self.file2 = None
        self.columns1 = []
        self.columns2 = []
        self.mappings = {}
        self.df1 = None
        self.df2 = None
        self.create_gui()

    def create_gui(self):
        """
        Create the GUI with buttons for file selection and column comparison.
        """
        tk.Button(self.root, text="Start Comparison", command=self.start_comparison).pack(pady=20)

    def start_comparison(self):
        """
        Start the comparison process by loading the first Excel file.
        """
        self.load_first_file()

    def load_first_file(self):
        """
        Load the first Excel file and read it into a DataFrame.
        """
        try:
            self.file1 = filedialog.askopenfilename(title="Select First Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
            if self.file1:
                self.df1 = pd.read_excel(self.file1)
                messagebox.showinfo("File Loaded", f"First file loaded successfully: {os.path.basename(self.file1)}")
                self.load_second_file()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load the first file: {e}")

    def load_second_file(self):
        """
        Load the second Excel file and read it into a DataFrame.
        """
        try:
            self.file2 = filedialog.askopenfilename(title="Select Second Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
            if self.file2:
                self.df2 = pd.read_excel(self.file2)
                messagebox.showinfo("File Loaded", f"Second file loaded successfully: {os.path.basename(self.file2)}")
                self.select_columns(self.df1, self.columns1, "Select Columns from First File", self.select_columns_from_second_file)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load the second file: {e}")

    def select_columns_from_second_file(self):
        """
        Trigger the column selection process for the second Excel file.
        """
        self.select_columns(self.df2, self.columns2, "Select Columns from Second File", self.validate_column_count)

    def select_columns(self, df, columns, title, next_step):
        """
        Prompt the user to select columns from the given DataFrame using a scrollable checkbox interface.

        :param df: The DataFrame to select the columns from.
        :param columns: The list to store the selected columns.
        :param title: The title for the selection window.
        :param next_step: The function to call after columns are selected.
        """
        top = Toplevel(self.root)
        top.title(title)
        scrollbar = Scrollbar(top, orient=VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox = Listbox(top, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set)
        for col in df.columns:
            listbox.insert(END, col)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        def on_submit():
            selected_indices = listbox.curselection()
            if not selected_indices:
                messagebox.showerror("Error", "Please select at least one column.")
                return
            for i in selected_indices:
                columns.append(df.columns[i])
            top.destroy()
            next_step()

        tk.Button(top, text="Submit", command=on_submit).pack(pady=10)

    def validate_column_count(self):
        """
        Validate that the number of selected columns from both files match and proceed to column mapping.
        """
        if len(self.columns1) != len(self.columns2):
            messagebox.showerror("Error", "The number of selected columns from both files must match.")
            self.columns1 = []
            self.columns2 = []
            self.start_comparison()
        else:
            self.map_columns(0)

    def map_columns(self, index):
        """
        Prompt the user to map columns between the two files.

        :param index: The current index of columns to map.
        """
        if index >= len(self.columns1):
            self.perform_comparison()
            return

        top = Toplevel(self.root)
        top.title(f"Map Columns {index+1}")

        tk.Label(top, text=f"Map '{self.columns1[index]}' to:").pack(pady=10)
        listbox = Listbox(top, selectmode=tk.SINGLE)
        for col in self.columns2:
            listbox.insert(END, col)
        listbox.pack(pady=10)

        def on_submit():
            selected_index = listbox.curselection()
            if not selected_index:
                messagebox.showerror("Error", "Please select a column to map to.")
                return
            self.mappings[self.columns1[index]] = self.columns2[selected_index[0]]
            top.destroy()
            self.map_columns(index + 1)

        tk.Button(top, text="Submit", command=on_submit).pack(pady=10)

    def perform_comparison(self):
        """
        Perform the comparison based on the mapped columns and generate the report.
        """
        try:
            # Prompt user to select save location and enforce .xlsx extension
            file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', title="Save Report As", filetypes=[("Excel files", "*.xlsx")])
            if not file_path:
                messagebox.showerror("Error", "No save location selected.")
                return

            # Define the maximum length for sheet names
            max_sheet_name_length = 31

            with pd.ExcelWriter(file_path) as writer:
                analysis = []
                for col1, col2 in self.mappings.items():
                    values1 = set(self.df1[col1])
                    values2 = set(self.df2[col2])

                    matching_values = values1.intersection(values2)
                    not_in_second = values1 - values2
                    not_in_first = values2 - values1

                    # Analysis summary
                    analysis.append({
                        'File': [os.path.basename(self.file1), os.path.basename(self.file2)],
                        'Column': [col1, col2],
                        'Total Values': [len(values1), len(values2)],
                        'Matching Values': len(matching_values),
                        'Not Matching Values': [len(not_in_second), len(not_in_first)]
                    })

                    # Not in second file
                    sheet_name1 = f'Not in {os.path.basename(self.file2)} ({col1})'
                    if len(sheet_name1) > max_sheet_name_length:
                        sheet_name1 = sheet_name1[:max_sheet_name_length]
                    pd.DataFrame(not_in_second, columns=[col1]).to_excel(writer, sheet_name=sheet_name1, index=False)

                    # Not in first file
                    sheet_name2 = f'Not in {os.path.basename(self.file1)} ({col2})'
                    if len(sheet_name2) > max_sheet_name_length:
                        sheet_name2 = sheet_name2[:max_sheet_name_length]
                    pd.DataFrame(not_in_first, columns=[col2]).to_excel(writer, sheet_name=sheet_name2, index=False)

                analysis_df = pd.concat([pd.DataFrame(a) for a in analysis])
                analysis_df.to_excel(writer, sheet_name='Analysis', index=False)

            messagebox.showinfo("Success", f"Comparison report generated successfully: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during comparison: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparator(root)
    root.mainloop()
