import tkinter as tk
from tkinter import filedialog, messagebox
import os

from excel_reader import read_input_excel
from processor import process_data

def run_app():
    input_file = filedialog.askopenfilename(
        title="Select Input Excel File (Project_Actual_Hours_for_Worker_by_Task_*.xlsx)",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    if not input_file:
        return

    output_dir = filedialog.askdirectory(title="Select Output Folder")
    if not output_dir:
        return

    try:
        df = read_input_excel(input_file)
        generated_files = process_data(df, output_dir)

        if not generated_files:
            messagebox.showwarning("No Output", "No files were generated (no valid rows/tasks).")
            return

        msg = f"Generated {len(generated_files)} file(s):\n\n"
        msg += "\n".join(f"- {os.path.basename(p)}" for p in generated_files)
        msg += f"\n\nOutput folder:\n{output_dir}"

        messagebox.showinfo("Success", msg)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def main():
    root = tk.Tk()
    root.title("Concentrix Timesheet Generator")
    root.geometry("600x220")

    btn = tk.Button(
        root,
        text="Generate Timesheet Files",
        font=("Arial", 13),
        width=35,
        height=2,
        command=run_app
    )
    btn.pack(expand=True)
    root.mainloop()

if __name__ == "__main__":
    main()
