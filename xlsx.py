import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
from tksheet import Sheet
import os

class TaskMeshApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Task Mesh - Personal Database")
        self.root.geometry("1200x750")
        self.root.configure(bg="#ffffff")

        # 1. Initial Data Generation (A001-A999)
        self.cols = ["block_id"] + [f"Column {i}" for i in range(2, 16)]
        self.initial_data = [[f"A{str(i).zfill(3)}"] + [""] * 14 for i in range(1, 1000)]

        # 2. Modern Header UI
        header = tk.Frame(root, bg="#ffffff", pady=15)
        header.pack(fill="x")
        
        tk.Label(header, text="TASK MANAGER", font=("Segoe UI", 16, "bold"), 
                 bg="#ffffff", fg="#222222").pack(side="left", padx=25)

        btn_config = {"font": ("Segoe UI", 9, "bold"), "relief": "flat", "padx": 15, "cursor": "hand2"}
        
        save_btn = tk.Button(header, text="SAVE AS .TXT", command=self.save_to_txt, 
                             bg="#000000", fg="white", **btn_config)
        save_btn.pack(side="right", padx=10)

        open_btn = tk.Button(header, text="OPEN FILE", command=self.open_any_file, 
                             bg="#666666", fg="white", **btn_config)
        open_btn.pack(side="right", padx=10)

        # 3. The Spreadsheet Grid
        self.sheet = Sheet(root, 
                           data=self.initial_data,
                           headers=self.cols,
                           theme="light blue")
        
        self.sheet.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        # 4. Enable Spreadsheet Features
        self.sheet.enable_bindings(
            "single_select", "edit_cell", "arrowkeys", 
            "copy", "cut", "paste", "delete", "undo", 
            "column_width_resize", "row_height_resize"
        )
        
        # Lock the ID column
        self.sheet.readonly_columns(columns=[0])

    def save_to_txt(self):
        try:
            data = self.sheet.get_sheet_data()
            df = pd.DataFrame(data, columns=self.cols)
            
            # Saving as a Tab-Separated .txt file
            file_path = os.path.join(os.getcwd(), "My_Task_Data.txt")
            df.to_csv(file_path, sep='\t', index=False)
            
            messagebox.showinfo("Saved", f"Data successfully saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {e}")

    def open_any_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt"), ("CSV Files", "*.csv"), ("Excel Files", "*.xlsx")]
        )
        
        if not file_path:
            return

        try:
            # Determine how to read the file
            if file_path.endswith('.txt'):
                df = pd.read_csv(file_path, sep='\t')
            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)

            # --- THE FIX: Replace NaN with empty strings ---
            df = df.fillna("")

            # Update the sheet
            self.sheet.set_sheet_data(df.values.tolist())
            self.sheet.headers(df.columns.tolist())
            
            # Make sure we reset the ID column to read-only for the new file
            self.sheet.readonly_columns(columns=[0])
            self.sheet.redraw()
            
            messagebox.showinfo("Loaded", f"Successfully loaded:\n{os.path.basename(file_path)}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Could not open this file.\nError: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TaskMeshApp(root)
    root.mainloop()
