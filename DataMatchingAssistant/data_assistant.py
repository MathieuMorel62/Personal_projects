import os
import re
import string
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from fuzzywuzzy import fuzz, process
import unidecode


class DataMatcherApp:
    def __init__(self, root: tk.Tk):
        """
        Initialize the application.

        Args:
            root (tk.Tk): The main window of the application.
        """
        self.root = root
        self.root.title("Data Matching Assistant")
        self.setup_ui()

        # Initialization of attributes
        self.start_time = None
        self.thread = None
        self.elapsed_time = 0
        self.selected_sheet1 = None
        self.selected_sheet2 = None
        self.update_sheets_and_columns()


    def setup_ui(self):
        """
        Set up the user interface of the application.
        """
        # Create a main frame to contain all UI elements.
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(expand=True, fill="both")

        # Configure custom styles for UI elements.
        style = ttk.Style()
        style.configure("Custom.TButton", padding=10, relief="raised")
        style.configure("Custom.TLabel", padding=5, font=("Helvetica", 24))
        style.configure("Custom.TEntry", padding=5, font=("Helvetica", 14))
        style.configure("Custom.Horizontal.TProgressbar", thickness=30)

        # Create a title label for the application.
        title_label = ttk.Label(main_frame, text="Data Matching Assistant", style="Custom.TLabel")
        title_label.pack(pady=10)

        # Create a frame for Excel file selection.
        file_frame = tk.LabelFrame(main_frame, text="1. Select the Excel file", padx=20, pady=20)
        file_frame.pack(padx=10, pady=10, fill="both")

        # Create a "Browse" button to select the file.
        self.btn_select_file = ttk.Button(file_frame, text="Browse", command=self.select_file, style="Custom.TButton")
        self.entry_file = ttk.Entry(file_frame, width=60, style="Custom.TEntry")
        self.btn_select_file.grid(row=0, column=0, padx=10, pady=10)
        self.entry_file.grid(row=0, column=1, padx=10, pady=10)

        # Create a frame for sheet selection.
        sheets_frame = tk.LabelFrame(main_frame, text="2. Select the 2 sheets", padx=20, pady=20)
        sheets_frame.pack(padx=10, pady=10, fill="both")

        # Create two dropdown lists for sheets and a "Validate sheets" button.
        self.sheet1_combobox = ttk.Combobox(sheets_frame, state="readonly", width=40, style="Custom.TCombobox")
        self.sheet2_combobox = ttk.Combobox(sheets_frame, state="readonly", width=40, style="Custom.TCombobox")
        self.btn_select_sheets = ttk.Button(sheets_frame, text="Validate sheets", command=self.select_sheets, style="Custom.TButton")
        
        # Place the dropdown lists and the button in the "sheets_frame".
        self.sheet1_combobox.grid(row=0, column=0, padx=10, pady=10)
        self.sheet2_combobox.grid(row=0, column=1, padx=10, pady=10)
        self.btn_select_sheets.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="n")

        # Create a frame for matching configuration.
        config_frame = tk.LabelFrame(main_frame, text="3. Matching configuration", padx=20, pady=20)
        config_frame.pack(padx=10, pady=10, fill="both")

        # Create a frame for file saving.
        save_frame = tk.LabelFrame(main_frame, text="4. Save the file", padx=20, pady=20)
        save_frame.pack(padx=10, pady=10, fill="both")

        # Create a text entry to specify the save location.
        self.entry_save_path = ttk.Entry(save_frame, width=60, style="Custom.TEntry")
        
        # Create a button to open a dialog box to select the save location.
        self.btn_save_path = ttk.Button(save_frame, text="Choose location", command=self.choose_save_path, style="Custom.TButton")
        self.entry_save_path.grid(row=0, column=1, padx=10, pady=10)
        self.btn_save_path.grid(row=0, column=0, padx=10, pady=10)

        # Create two dropdown lists to select columns to compare.
        self.column1_combobox = ttk.Combobox(config_frame, state="readonly", width=40, style="Custom.TCombobox")
        self.column2_combobox = ttk.Combobox(config_frame, state="readonly", width=40, style="Custom.TCombobox")
        
        # Create a label to indicate the matching threshold.
        self.label_threshold = ttk.Label(config_frame, text="Matching threshold :", style="Custom.TLabel")
        
        # Create a text entry to specify the matching threshold (default to 50).
        self.entry_threshold = ttk.Entry(config_frame, width=5, style="Custom.TEntry")
        self.entry_threshold.insert(0, "50")

        # Place the dropdown lists, label, and text entry in the "config_frame".
        self.column1_combobox.grid(row=0, column=0, padx=10, pady=10)
        self.column2_combobox.grid(row=0, column=1, padx=10, pady=10)
        self.label_threshold.grid(row=1, column=0, padx=10, pady=10)
        self.entry_threshold.grid(row=1, column=1, padx=10, pady=10)

        # Create a frame for actions (script execution).
        action_frame = tk.Frame(main_frame, padx=20, pady=20)
        action_frame.pack(padx=10, pady=10, fill="both")

        # Create a "Run" button to start the matching process.
        self.btn_run = ttk.Button(action_frame, text="Run", command=self.run_script, style="Custom.TButton")

        # Create a progress bar to track the processing status.
        self.progress = ttk.Progressbar(action_frame, orient="horizontal", length=700, mode="determinate", style="Custom.Horizontal.TProgressbar")

        # Create a label to display processing time.
        self.label_elapsed_time = ttk.Label(action_frame, text="Processing time: 0.00 seconds", font=("Helvetica", 22), style="Custom.TLabel")

        # Place the button, progress bar, and label in the "action_frame".
        self.btn_run.pack(padx=10, pady=20)
        self.progress.pack(padx=10, pady=10)
        self.label_elapsed_time.pack(padx=10, pady=10)


    def select_file(self):
        """
        Open a dialog box to select an Excel file.
        """
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, file_path)
            self.update_sheets_and_columns()


    def validate_file_exists(self, file_path: str) -> bool:
        """
        Validate the existence of a file.

        Args:
            file_path (str): The path of the file to validate.

        Returns:
            bool: True if the file exists, False otherwise.
        """
        if os.path.exists(file_path) and os.path.isfile(file_path):
            return True
        else:
            messagebox.showerror("Error", "Please select a file.")
            return False


    def preprocess_text(self, text):
        """
        Preprocess the text by converting it to lowercase, removing punctuation,
        normalizing accents, and removing excess spaces.

        Args:
            text (str): The text to preprocess.

        Returns:
            str: The preprocessed text.
        """
        if isinstance(text, str):
            text = text.lower()
            text = re.sub(r'[{}]'.format(string.punctuation), ' ', text)
            text = re.sub(r'\s+', ' ', text).strip()
            text = unidecode.unidecode(text)
        return text


    def find_best_match_with_score(self, row, column_name, pnt_data, match_column, match_indicator):
        """
        Search for the best match in pnt_data for a given row.

        Args:
            row (pd.Series): The row for which we are searching for a match.
            column_name (str): The name of the row column to compare.
            pnt_data (pd.DataFrame): The data in which to search for the match.
            match_column (str): The name of the column in pnt_data to compare.
            match_indicator (int): The minimum matching threshold.

        Returns:
            Tuple[str, int]: A tuple containing the best match (None if none) and the match score (None if none).
        """
        matches = process.extract(row[column_name], pnt_data[match_column], limit=1, scorer=fuzz.token_set_ratio)
        if matches and matches[0][1] > match_indicator:
            return (matches[0][0], matches[0][1])
        return (None, None)


    def update_progress(self, progress):
        """
        Update the progress bar.

        Args:
            progress (float): The progress value (0-100).
        """
        self.progress['value'] = progress
        self.root.update_idletasks()


    def update_sheets_and_columns(self):
        """
        Update the dropdown lists of sheets and available columns in the selected file.
        """
        file_path = self.entry_file.get()
        if file_path:
            with pd.ExcelFile(file_path) as xls:
                sheets = xls.sheet_names
                self.sheet1_combobox['values'] = sheets
                self.sheet2_combobox['values'] = sheets

            self.sheet1_combobox['state'] = 'readonly'
            self.sheet2_combobox['state'] = 'readonly'


    def select_sheets(self):
        """
        Validate and save the selected sheets.
        Also update the available columns.
        """
        self.selected_sheet1 = self.sheet1_combobox.get()
        self.selected_sheet2 = self.sheet2_combobox.get()
        if self.selected_sheet1 and self.selected_sheet2:
            messagebox.showinfo("Success", f"Selected sheets : {self.selected_sheet1}, {self.selected_sheet2}")
            self.column1_combobox['values'] = pd.read_excel(self.entry_file.get(), sheet_name=self.selected_sheet1).columns.tolist()
            self.column2_combobox['values'] = pd.read_excel(self.entry_file.get(), sheet_name=self.selected_sheet2).columns.tolist()


    def choose_save_path(self):
        """
        Open a dialog box to choose the save location.
        """
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            self.entry_save_path.delete(0, tk.END)
            self.entry_save_path.insert(0, save_path)


    def perform_matching(self):
        """
        Perform data matching.
        """
        file_path = self.entry_file.get()
        MATCH_INDICATOR = int(self.entry_threshold.get())
        output_file = self.entry_save_path.get() or "Final_File.xlsx"
        sheet1 = self.selected_sheet1
        sheet2 = self.selected_sheet2
        column1 = self.column1_combobox.get()
        column2 = self.column2_combobox.get()

        if not self.validate_sheet_and_columns(file_path, sheet1, sheet2, column1, column2):
            return

        try:
            df1 = pd.read_excel(file_path, sheet_name=sheet1)
            df2 = pd.read_excel(file_path, sheet_name=sheet2)

            df1[column1] = df1[column1].apply(self.preprocess_text)
            df2[column2] = df2[column2].apply(self.preprocess_text)

            total_rows = len(df1)
            for index, row in df1.iterrows():
                best_match, match_score = self.find_best_match_with_score(row, column1, df2, column2, MATCH_INDICATOR)
                df1.at[index, 'Best Match'] = best_match
                df1.at[index, 'Match Percentage'] = match_score
                progress = (index + 1) / total_rows * 100
                self.update_progress(progress)

                est_time_seconds = total_rows * 2
                elapsed_seconds = max(0, est_time_seconds - (time.time() - self.start_time))
                hours, remainder = divmod(elapsed_seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                time_str = f"{int(hours)}h {int(minutes)}min {int(seconds)}sec"
                self.label_elapsed_time.config(text=f"Processing time : {time_str}")
                self.root.update_idletasks()

            df1.to_excel(output_file, index=False)
            elapsed_time = time.time() - self.start_time
            self.label_elapsed_time.config(text=f"Total processing time : {elapsed_time:.2f} secondes")
            self.root.update_idletasks()
            messagebox.showinfo("Success", f"Matching completed in {elapsed_time:.2f} seconds. Result saved in {output_file}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred : {str(e)}")


    def validate_threshold(self):
        """
        Validate the matching threshold entered by the user.

        Returns:
            bool: True if the threshold is valid, False otherwise.
        """
        try:
            threshold = int(self.entry_threshold.get())
            if 0 <= threshold <= 100:
                return True
            else:
                messagebox.showerror("Error", "The matching threshold must be a number between 0 and 100.")
                return False
        except ValueError:
            messagebox.showerror("Error", "The matching threshold must be a valid number.")
            return False


    def validate_sheet_and_columns(self, file_path, sheet1, sheet2, column1, column2):
        """
        Validate the selected sheets and columns.

        Args:
            file_path (str): The path of the Excel file.
            sheet1 (str): The name of the first sheet.
            sheet2 (str): The name of the second sheet.
            column1 (str): The name of the column of the first sheet.
            column2 (str): The name of the column of the second sheet.

        Returns:
            bool: True if the sheets and columns are valid, False otherwise.
        """
        try:
            with pd.ExcelFile(file_path) as xls:
                sheet_names = xls.sheet_names
                if sheet1 not in sheet_names or sheet2 not in sheet_names:
                    messagebox.showerror("Error", "One of the selected sheets does not exist in the file.")
                    return False

                df1 = pd.read_excel(file_path, sheet_name=sheet1)
                df2 = pd.read_excel(file_path, sheet_name=sheet2)

                if column1 not in df1.columns.tolist() or column2 not in df2.columns.tolist():
                    messagebox.showerror("Error", "One of the selected columns does not exist in the corresponding sheet.")
                    return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred : {str(e)}")
            return False


    def run_script(self):
        """
        Execute the data matching script.
        """
        if self.thread and self.thread.is_alive():
            messagebox.showerror("Error", "Processing is already in progress.")
            return

        file_path = self.entry_file.get()
        save_path = self.entry_save_path.get()

        if not self.validate_file_exists(file_path) or not self.validate_threshold():
            return

        if not self.validate_sheet_and_columns(file_path, self.selected_sheet1, self.selected_sheet2,
                                               self.column1_combobox.get(), self.column2_combobox.get()):
            return

        self.start_time = time.time()
        self.thread = threading.Thread(target=self.perform_matching)
        self.thread.start()


if __name__ == "__main__":
    root = tk.Tk()
    app = DataMatcherApp(root)
    root.mainloop()
