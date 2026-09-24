import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import APP_TITLE
from .processor import process_files


class EmployeeCleanupApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.state("zoomed")

        self.employee_file = tk.StringVar()
        self.records_file = tk.StringVar()
        self.output_file = tk.StringVar()

        self.employee_sheet = tk.StringVar(value="Sheet1")
        self.records_sheet = tk.StringVar(value="Profund")

        self.employee_id_column = tk.StringVar(value="Employee ID")
        self.employee_status = tk.StringVar(value="Employee Status")
        self.records_employee_id_column = tk.StringVar(value="Emp_ID")
        self.records_id = tk.StringVar(value="Id")

        self.status = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=20)
        root.pack(fill="both", expand=True)

        ttk.Label(
            root,
            text=APP_TITLE,
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w")

        ttk.Label(
            root,
            text=(
                "Find terminated employees, remove their older duplicate records, "
                "and blank the final record."
            ),
            wraplength=820,
        ).pack(anchor="w", pady=(5, 18))

        files = ttk.LabelFrame(root, text="Excel files", padding=12)
        files.pack(fill="x", pady=6)

        self._file_row(files, "Employee master", self.employee_file, self.select_employee)
        self._file_row(files, "Employee records", self.records_file, self.select_records)
        self._file_row(files, "Output file", self.output_file, self.select_output)

        config = ttk.LabelFrame(root, text="Configuration", padding=12)
        config.pack(fill="x", pady=10)

        fields = [
            ("Employee sheet", self.employee_sheet),
            ("Records sheet", self.records_sheet),
            ("Employee ID column", self.employee_id_column),
            ("Employee status column", self.employee_status),
            ("Records employee ID column", self.records_employee_id_column),
            ("Records Id column to keep", self.records_id),
        ]

        for i, (label, variable) in enumerate(fields):
            row = i // 2
            col = (i % 2) * 2

            ttk.Label(config, text=label).grid(
                row=row, column=col, sticky="w", padx=(0, 8), pady=5
            )
            ttk.Entry(config, textvariable=variable, width=35).grid(
                row=row,
                column=col + 1,
                sticky="ew",
                padx=(0, 18),
                pady=5,
            )

        config.columnconfigure(1, weight=1)
        config.columnconfigure(3, weight=1)

        rules = ttk.LabelFrame(root, text="Processing rules", padding=12)
        rules.pack(fill="x", pady=6)

        ttk.Label(
            rules,
            text=(
                "1. Select employees whose Employee Status is Terminated.\n"
                "2. Match Employee ID to the records Emp_ID column.\n"
                "3. For each terminated employee, keep only the physically last matching row.\n"
                "4. Delete earlier matching rows.\n"
                "5. Keep only the records Id value in the retained row.\n"
                "6. Add a Processing Log sheet.\n"
                "7. Never modify the original input files."
            ),
            justify="left",
        ).pack(anchor="w")

        buttons = ttk.Frame(root)
        buttons.pack(fill="x", pady=18)

        self.process_button = ttk.Button(
            buttons,
            text="PROCESS FILES",
            command=self.process,
        )
        self.process_button.pack(side="left")

        ttk.Button(
            buttons,
            text="Reset",
            command=self.reset,
        ).pack(side="left", padx=10)

        ttk.Label(
            root,
            textvariable=self.status,
            wraplength=820,
        ).pack(anchor="w")

    def _file_row(self, parent, label, variable, command):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=5)

        ttk.Label(row, text=label, width=20).pack(side="left")

        ttk.Entry(
            row,
            textvariable=variable,
        ).pack(side="left", fill="x", expand=True, padx=8)

        ttk.Button(
            row,
            text="Browse...",
            command=command,
        ).pack(side="right")

    def select_employee(self):
        path = filedialog.askopenfilename(
            title="Select employee master file",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.employee_file.set(path)

    def select_records(self):
        path = filedialog.askopenfilename(
            title="Select employee records file",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.records_file.set(path)

    def select_output(self):
        path = filedialog.asksaveasfilename(
            title="Save cleaned Excel file",
            defaultextension=".xlsx",
            initialfile="employee_records_cleaned.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.output_file.set(path)

    def reset(self):
        self.employee_file.set("")
        self.records_file.set("")
        self.output_file.set("")
        self.status.set("Ready")

    def process(self):
        try:
            if not self.employee_file.get():
                raise ValueError("Select the employee master Excel file.")
            if not self.records_file.get():
                raise ValueError("Select the employee records Excel file.")

            output = self.output_file.get()

            if not output:
                default_name = (
                    Path(self.records_file.get()).stem
                    + "_cleaned.xlsx"
                )
                output = filedialog.asksaveasfilename(
                    title="Save cleaned Excel file",
                    defaultextension=".xlsx",
                    initialfile=default_name,
                    filetypes=[("Excel files", "*.xlsx")],
                )

                if not output:
                    return

                self.output_file.set(output)

            self.process_button.config(state="disabled")
            self.status.set("Processing...")

            self.update_idletasks()

            result = process_files(
                employee_file=self.employee_file.get(),
                records_file=self.records_file.get(),
                output_file=output,
                employee_sheet=self.employee_sheet.get(),
                records_sheet=self.records_sheet.get(),
                employee_id_column=self.employee_id_column.get(),
                employee_status=self.employee_status.get(),
                records_employee_id_column=self.records_employee_id_column.get(),
                records_id=self.records_id.get(),
            )

            self.status.set("Processing completed successfully.")

            messagebox.showinfo(
                "Completed",
                "Processing completed successfully.\n\n"
                f"Terminated employees: {result['terminated']}\n"
                f"Found in records: {result['matched']}\n"
                f"Records deleted: {result['deleted']}\n"
                f"Records retained and blanked: {result['blanked']}\n"
                f"Not found in records: {result['not_found']}\n\n"
                f"Output:\n{output}",
            )

        except Exception as exc:
            self.status.set("Processing failed.")
            messagebox.showerror(
                "Processing failed",
                str(exc),
            )

        finally:
            self.process_button.config(state="normal")

