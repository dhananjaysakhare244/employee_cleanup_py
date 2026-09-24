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

        self.head_count_path = tk.StringVar()
        self.classic_user_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.head_count_sheet_name = tk.StringVar(value="Sheet1")
        self.classic_user_sheet_name = tk.StringVar(value="Profund")

        self.head_count_employee_id_header = tk.StringVar(value="Employee ID")
        self.head_count_status_header = tk.StringVar(value="Employee Status")
        self.classic_user_employee_id_header = tk.StringVar(value="Emp_ID")
        self.classic_user_id_header = tk.StringVar(value="Id")

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
                "Find terminated employees in Head Count and clean matching records "
                "in Classic User."
            ),
            wraplength=820,
        ).pack(anchor="w", pady=(5, 18))

        files = ttk.LabelFrame(root, text="Workbooks", padding=12)
        files.pack(fill="x", pady=6)

        self._file_row(
            files, "Head Count workbook", self.head_count_path, self.select_head_count
        )
        self._file_row(
            files, "Classic User workbook", self.classic_user_path, self.select_classic_user
        )
        self._file_row(files, "Cleaned workbook", self.output_path, self.select_output)

        config = ttk.LabelFrame(root, text="Configuration", padding=12)
        config.pack(fill="x", pady=10)

        fields = [
            ("Head Count worksheet", self.head_count_sheet_name),
            ("Classic User worksheet", self.classic_user_sheet_name),
            ("Head Count employee ID header", self.head_count_employee_id_header),
            ("Head Count status header", self.head_count_status_header),
            ("Classic User employee ID header", self.classic_user_employee_id_header),
            ("Classic User ID header to keep", self.classic_user_id_header),
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
                "1. Select Head Count employees with status Terminated.\n"
                "2. Match Head Count Employee ID to Classic User Emp_ID.\n"
                "3. Keep the last matching Classic User row and delete earlier matches.\n"
                "4. Keep its Id and clear the other values.\n"
                "5. Add a Processing Log worksheet. Original workbooks are not changed."
            ),
            justify="left",
        ).pack(anchor="w")

        buttons = ttk.Frame(root)
        buttons.pack(fill="x", pady=18)

        self.process_button = ttk.Button(
            buttons,
            text="Process Workbooks",
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

    def select_head_count(self):
        path = filedialog.askopenfilename(
            title="Select the Head Count workbook",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.head_count_path.set(path)

    def select_classic_user(self):
        path = filedialog.askopenfilename(
            title="Select the Classic User workbook",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.classic_user_path.set(path)

    def select_output(self):
        path = filedialog.asksaveasfilename(
            title="Save cleaned Classic User workbook",
            defaultextension=".xlsx",
            initialfile="classic_user_cleaned.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def reset(self):
        self.head_count_path.set("")
        self.classic_user_path.set("")
        self.output_path.set("")
        self.status.set("Ready")

    def process(self):
        try:
            if not self.head_count_path.get():
                raise ValueError("Choose the Head Count workbook first.")
            if not self.classic_user_path.get():
                raise ValueError("Choose the Classic User workbook first.")

            output = self.output_path.get()

            if not output:
                default_name = (
                    Path(self.classic_user_path.get()).stem
                    + "_cleaned.xlsx"
                )
                output = filedialog.asksaveasfilename(
                    title="Save cleaned Classic User workbook",
                    defaultextension=".xlsx",
                    initialfile=default_name,
                    filetypes=[("Excel files", "*.xlsx")],
                )

                if not output:
                    return

                self.output_path.set(output)

            self.process_button.config(state="disabled")
            self.status.set("Processing...")

            self.update_idletasks()

            result = process_files(
                head_count_path=self.head_count_path.get(),
                classic_user_path=self.classic_user_path.get(),
                output_path=output,
                head_count_sheet_name=self.head_count_sheet_name.get(),
                classic_user_sheet_name=self.classic_user_sheet_name.get(),
                head_count_employee_id_header=self.head_count_employee_id_header.get(),
                head_count_status_header=self.head_count_status_header.get(),
                classic_user_employee_id_header=self.classic_user_employee_id_header.get(),
                classic_user_id_header=self.classic_user_id_header.get(),
            )

            self.status.set("Processing completed successfully.")

            messagebox.showinfo(
                "Completed",
                "Processing completed successfully.\n\n"
                f"Terminated employees: {result['terminated']}\n"
                f"Found in Classic User: {result['matched']}\n"
                f"Classic User rows deleted: {result['deleted']}\n"
                f"Classic User rows retained and cleared: {result['blanked']}\n"
                f"Not found in Classic User: {result['not_found']}\n\n"
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

