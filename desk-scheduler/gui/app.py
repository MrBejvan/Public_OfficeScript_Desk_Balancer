"""
Tkinter desktop UI for the Desk Coverage Scheduler.

Kept intentionally simple (no external GUI toolkits) so the whole app runs
from the Python standard library on a locked-down, offline work laptop:
select a file, pick a period, run the balancer, review results, export a
calendar. Every action that touches the filesystem or parses user data is
wrapped so a bad file or bad date shows a plain-English message instead of
a crash or a stack trace.
"""

from __future__ import annotations

import logging
import tempfile
import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from core.exporter import export_schedule
from core.importer import KronosImportError, load_shift_requests
from core.models import Desk, SchedulePeriod
from core.scheduler import build_schedule

logging.basicConfig(
    filename=Path(__file__).resolve().parent.parent / "app.log",
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger("desk_scheduler")

DATE_FORMAT = "%m/%d/%Y"


class DeskSchedulerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Desk Coverage Scheduler")
        self.geometry("880x640")
        self.minsize(760, 560)

        self.selected_file: Path | None = None
        self.current_period: SchedulePeriod | None = None

        self._build_menu()
        self._build_layout()

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------
    def _build_menu(self) -> None:
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.config(menu=menubar)

    def _build_layout(self) -> None:
        pad = {"padx": 8, "pady": 6}

        step1 = ttk.LabelFrame(self, text="Step 1 - Choose the Kronos export")
        step1.pack(fill="x", **pad)

        self.file_var = tk.StringVar(value="No file selected")
        ttk.Label(step1, textvariable=self.file_var, foreground="#444").pack(
            side="left", padx=8, pady=8, fill="x", expand=True
        )
        ttk.Button(step1, text="Browse...", command=self._browse_file).pack(
            side="right", padx=8, pady=8
        )
        ttk.Button(
            step1, text="Use Sample Data", command=self._use_sample_data
        ).pack(side="right", padx=4, pady=8)

        step2 = ttk.LabelFrame(self, text="Step 2 - Schedule period")
        step2.pack(fill="x", **pad)

        today = date.today()
        first_of_month = today.replace(day=1)
        next_month = (first_of_month.month % 12) + 1
        next_year = first_of_month.year + (1 if first_of_month.month == 12 else 0)
        last_of_month = date(next_year, next_month, 1)
        last_of_month = last_of_month.fromordinal(last_of_month.toordinal() - 1)

        ttk.Label(step2, text="Start date (MM/DD/YYYY):").grid(
            row=0, column=0, padx=8, pady=8, sticky="w"
        )
        self.start_var = tk.StringVar(value=first_of_month.strftime(DATE_FORMAT))
        ttk.Entry(step2, textvariable=self.start_var, width=14).grid(
            row=0, column=1, padx=4, pady=8
        )

        ttk.Label(step2, text="End date (MM/DD/YYYY):").grid(
            row=0, column=2, padx=8, pady=8, sticky="w"
        )
        self.end_var = tk.StringVar(value=last_of_month.strftime(DATE_FORMAT))
        ttk.Entry(step2, textvariable=self.end_var, width=14).grid(
            row=0, column=3, padx=4, pady=8
        )

        step3 = ttk.LabelFrame(self, text="Step 3 - Balance and export")
        step3.pack(fill="x", **pad)
        ttk.Button(step3, text="Run Balance", command=self._run_balance).pack(
            side="left", padx=8, pady=8
        )
        self.export_button = ttk.Button(
            step3,
            text="Export Calendar...",
            command=self._export_calendar,
            state="disabled",
        )
        self.export_button.pack(side="left", padx=4, pady=8)

        self.summary_var = tk.StringVar(value="Run a balance to see results here.")
        ttk.Label(step3, textvariable=self.summary_var, foreground="#444").pack(
            side="left", padx=16, pady=8
        )

        results = ttk.LabelFrame(self, text="Preview (by day)")
        results.pack(fill="both", expand=True, **pad)

        columns = ("date", "east", "west", "vab", "resource")
        self.tree = ttk.Treeview(results, columns=columns, show="headings", height=10)
        headings = {
            "date": "Date",
            "east": "East",
            "west": "West",
            "vab": "VAB",
            "resource": "Resource",
        }
        for col, text in headings.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=140 if col == "date" else 150, anchor="center")
        self.tree.tag_configure("shortfall", background="#FCE4E4")
        scroll = ttk.Scrollbar(results, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        issues_frame = ttk.LabelFrame(self, text="Import notes / items needing attention")
        issues_frame.pack(fill="both", expand=False, **pad)
        self.issues_text = tk.Text(issues_frame, height=6, wrap="word", state="disabled")
        self.issues_text.pack(fill="both", expand=True, padx=4, pady=4)

    def _show_about(self) -> None:
        messagebox.showinfo(
            "About",
            "Desk Coverage Scheduler\n\n"
            "Balances submitted shift requests across the East, West, VAB, "
            "and Resource desks and produces a printable monthly calendar.\n\n"
            "Runs entirely offline; no data ever leaves this computer.",
        )

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _browse_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Kronos export",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        self.selected_file = Path(path)
        self.file_var.set(str(self.selected_file))

    def _use_sample_data(self) -> None:
        try:
            from sample_data.generate_sample import generate_sample_workbook

            start = self._parse_date(self.start_var.get(), "start")
            end = self._parse_date(self.end_var.get(), "end")
            tmp_path = Path(tempfile.gettempdir()) / "desk_scheduler_sample.xlsx"
            generate_sample_workbook(tmp_path, start, end)
            self.selected_file = tmp_path
            self.file_var.set(f"{tmp_path}  (sample data)")
        except Exception:
            logger.exception("Failed to generate sample data")
            messagebox.showerror(
                "Sample Data Error", "Could not generate sample data. See app.log."
            )

    def _parse_date(self, text: str, label: str) -> date:
        text = text.strip()
        try:
            return datetime.strptime(text, DATE_FORMAT).date()
        except ValueError:
            raise ValueError(
                f"The {label} date '{text}' isn't valid. Please use MM/DD/YYYY."
            )

    def _run_balance(self) -> None:
        if self.selected_file is None:
            messagebox.showwarning(
                "No file selected", "Please choose a Kronos export file first."
            )
            return

        try:
            start = self._parse_date(self.start_var.get(), "start")
            end = self._parse_date(self.end_var.get(), "end")
            if start > end:
                raise ValueError("The start date must be on or before the end date.")
        except ValueError as exc:
            messagebox.showerror("Invalid Date", str(exc))
            return

        try:
            result = load_shift_requests(self.selected_file)
        except KronosImportError as exc:
            messagebox.showerror("Import Error", str(exc))
            return
        except Exception:
            logger.exception("Unexpected error while importing file")
            messagebox.showerror(
                "Import Error",
                "Something went wrong reading this file. See app.log for details.",
            )
            return

        if not result.requests:
            messagebox.showwarning(
                "No Usable Data",
                "No valid shift requests were found in this file. "
                "Check the import notes panel for details.",
            )

        try:
            period = build_schedule(result.requests, start, end)
        except Exception:
            logger.exception("Unexpected error while balancing schedule")
            messagebox.showerror(
                "Scheduling Error",
                "Something went wrong building the schedule. See app.log for details.",
            )
            return

        self.current_period = period
        self._populate_results(period, result.issues)
        self.export_button.config(state="normal" if period.assignments else "disabled")

    def _populate_results(self, period: SchedulePeriod, issues) -> None:
        for row in self.tree.get_children():
            self.tree.delete(row)

        day = period.start
        from datetime import timedelta

        while day <= period.end:
            by_desk = {a.desk: a.employee for a in period.assignments_for_day(day)}
            has_open = False
            row_values = [day.strftime("%a %m/%d/%Y")]
            for desk in Desk.priority_order():
                emp = by_desk.get(desk)
                if desk not in by_desk:
                    row_values.append("")
                elif emp is None:
                    row_values.append("OPEN")
                    has_open = True
                else:
                    row_values.append(emp.name)
            tags = ("shortfall",) if has_open else ()
            self.tree.insert("", "end", values=row_values, tags=tags)
            day += timedelta(days=1)

        self.issues_text.config(state="normal")
        self.issues_text.delete("1.0", "end")
        lines = []
        for issue in issues:
            lines.append(f"Row {issue.row_number}: {issue.message}")
        if period.shortfalls:
            lines.append("")
            lines.append(f"{len(period.shortfalls)} desk-day(s) could not be filled:")
            for flag in period.shortfalls:
                lines.append(
                    f"  {flag.work_date.strftime('%m/%d/%Y')} - {flag.desk.value}: {flag.reason}"
                )
        if not lines:
            lines = ["No issues found. All requests imported and all desks covered."]
        self.issues_text.insert("1.0", "\n".join(lines))
        self.issues_text.config(state="disabled")

        filled = sum(1 for a in period.assignments if a.employee is not None)
        total = len(period.assignments)
        self.summary_var.set(
            f"{filled}/{total} desk-shifts filled, {len(period.shortfalls)} open, "
            f"{len(issues)} import note(s)."
        )

    def _export_calendar(self) -> None:
        if self.current_period is None:
            return
        path = filedialog.asksaveasfilename(
            title="Save calendar as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="desk_coverage_calendar.xlsx",
        )
        if not path:
            return
        try:
            export_schedule(self.current_period, path)
        except Exception:
            logger.exception("Unexpected error while exporting calendar")
            messagebox.showerror(
                "Export Error",
                "Something went wrong saving the calendar. See app.log for details.",
            )
            return
        messagebox.showinfo("Export Complete", f"Calendar saved to:\n{path}")


def main() -> None:
    app = DeskSchedulerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
