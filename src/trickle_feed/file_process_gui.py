
import tkinter as tk
from tkinter import filedialog
import os
from main import process_file
from constants import UTILITY_NAME, UTILITY_VERSION, COPYRIGHT_CLAUSE

class FileProcess(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent=parent
        self.src_path = ""
        self.dst_path = ""
        self.active_colour = "blue"
        self.inactive_colour = "red"
        self.init_ui()

    def init_ui(self):
        self.parent.title(f"{UTILITY_NAME} {UTILITY_VERSION}")
        self.parent.resizable(False, False)

        self.add_path_panel("Source Directory", self.src_field())
        self.add_path_panel("Report Directory", self.dst_field())

        result_panel = tk.Frame(self)
        tk.Label(result_panel, text="Result:").pack(side=tk.LEFT)
        self.result_field = tk.Entry(result_panel, width=50, font=("Arial", 10))
        self.result_field.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        result_panel.pack(fill=tk.X, padx=10, pady=5)

        btn_panel = tk.Frame(self)
        tk.Button(btn_panel, text="Clear Result", command=self.clear_result).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_panel, text="Generate Report", command=self.generate_report,
                  bg="lightgreen", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        btn_panel.pack(pady=10)

        copyright_panel = tk.Frame(self)
        tk.Label(copyright_panel, text=COPYRIGHT_CLAUSE, font=("Arial", 8), fg="gray").pack()
        copyright_panel.pack(pady=5)

    def src_field(self):
        panel = tk.Frame(self)
        tk.Button(panel, text="Source Directory",
                  command=lambda: self.browse_folder(self.src_entry)).pack(side=tk.LEFT, padx=5)
        self.src_entry = tk.Entry(panel, width=50)
        self.src_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        return panel

    def dst_field(self):
        panel = tk.Frame(self)
        tk.Button(panel, text="Report Directory",
                  command=lambda: self.browse_folder(self.dst_entry)).pack(side=tk.LEFT, padx=5)
        self.dst_entry = tk.Entry(panel, width=50)
        self.dst_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        return panel

    def add_path_panel(self, panel, field_panel):
        field_panel.pack(fill=tk.X, padx=10, pady=5)

    def browse_folder(self, entry):
        folder = filedialog.askdirectory(initialdir=os.getcwd())
        if folder:
            entry.delete(0, tk.END)
            entry.insert(0, folder)

    def clear_result(self):
        self.result_field.delete(0, tk.END)

    def generate_report(self):
        self.src_path = self.src_entry.get().strip()
        self.dst_path = self.dst_entry.get().strip()
        self.result_field.delete(0, tk.END)
        message = ""
        err_count = 0
        
        if self.src_path and self.dst_path:
            try:
                report = process_file([self.src_path, self.dst_path])
                if not report.isErrorFlag():
                    message = f"{report.getReportName()} file is generated!"
                else:
                    message = report.getErrorCode()
                    err_count += 1
            except Exception as ex:
                message = "THERE IS ISSUE IN FILE PROCESSING."
                err_count += 1
        else:
            message = "Please enter proper "
            if not self.src_path:
                message += "Source Directory "
                err_count += 1
            if not self.dst_path:
                if err_count > 0: message += "and "
                message += "Destination Directory"
                err_count += 1
            message += " path."

        self.result_field.config(fg=self.inactive_colour if err_count > 0 else self.active_colour,
                                 bg="lightpink" if err_count > 0 else "lightgreen",
                                 font=("Arial", 12, "bold"))
        self.result_field.insert(0, message)
        print(f"---------------OUTPUT MSG: {message}")

def main():
    root = tk.Tk()
    app = FileProcess(root)
    app.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    root.update()
    root.geometry("800x300+300+200")
    root.mainloop()

if __name__ == "__main__":
    main()
