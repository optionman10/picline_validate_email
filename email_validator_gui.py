"""
email_validator_gui.py
----------------------
Desktop tool (tkinter) to validate email addresses in an Excel file.

Features:
  * Browse for an .xlsx/.xlsm file
  * Auto-detects the sheet + column that contains email addresses
  * Validates syntax, naming conventions, disposable domains,
    common typos, and DNS MX records (background thread + progress bar)
  * Summary counts: VALID / VALID WITH WARNINGS / INVALID
  * Scrollable results table with per-row error description
  * Exports '<name>_validated.xlsx' with Status / Error Description /
    Suggestion columns; invalid rows highlighted red, warnings yellow.

Run:  python email_validator_gui.py
"""

from __future__ import annotations

import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from email_validator import (
    DomainChecker,
    Status,
    ValidationResult,
    validate_email,
)

EMAIL_HEADER_HINTS = {"email", "e-mail", "email address", "e-mail address",
                      "mail", "emailaddress", "emails"}

FILL_ERROR = PatternFill("solid", start_color="FFC7CE")    # light red
FILL_WARNING = PatternFill("solid", start_color="FFEB9C")  # light yellow
FONT_ERROR = Font(color="9C0006")
FONT_WARNING = Font(color="9C6500")

ROW_TAG_COLORS = {
    Status.INVALID: ("#fdecec", "#9C0006"),
    Status.VALID_WITH_WARNINGS: ("#fffbe6", "#8a6d00"),
    Status.VALID: ("#f0fff0", "#1e7d1e"),
}

FILTER_OPTIONS = ("All", "Valid", "Warnings", "Invalid")


# --------------------------------------------------------------------------
# Excel helpers
# --------------------------------------------------------------------------

def detect_email_sheet_column(path: Path) -> tuple[str, int, int] | None:
    """Return (sheet_name, header_row, column_index 1-based) of the email
    column, or None if nothing email-shaped was found.

    Strategy: look for a header containing 'email'/'mail' in the first
    10 rows of any sheet; fall back to the column with the most values
    containing exactly one '@'.
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        best_scan: tuple[str, int, int, int] | None = None  # sheet,row,col,score
        for ws in wb.worksheets:
            rows = ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 60),
                                values_only=True)
            rows = list(rows)
            if not rows:
                continue
            # 1) header hint
            for r_idx, row in enumerate(rows[:10], start=1):
                for c_idx, cell in enumerate(row, start=1):
                    if isinstance(cell, str) and \
                       cell.strip().lower() in EMAIL_HEADER_HINTS:
                        return ws.title, r_idx, c_idx
            # 2) scan: column with most '@' values
            scores: dict[int, int] = {}
            for row in rows:
                for c_idx, cell in enumerate(row, start=1):
                    if isinstance(cell, str) and cell.count("@") == 1:
                        scores[c_idx] = scores.get(c_idx, 0) + 1
            if scores:
                col, score = max(scores.items(), key=lambda kv: kv[1])
                if best_scan is None or score > best_scan[3]:
                    # header row = first row above the first email value
                    header_row = 1
                    for r_idx, row in enumerate(rows, start=1):
                        if (len(row) >= col and isinstance(row[col - 1], str)
                                and row[col - 1].count("@") == 1):
                            header_row = max(1, r_idx - 1)
                            break
                    best_scan = (ws.title, header_row, col, score)
        if best_scan:
            return best_scan[0], best_scan[1], best_scan[2]
        return None
    finally:
        wb.close()


def load_emails(path: Path, sheet: str, header_row: int, col: int) -> list[tuple[int, str]]:
    """Return list of (excel_row_number, raw_value) below the header."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        out = []
        for row in ws.iter_rows(min_row=header_row + 1, min_col=col, max_col=col):
            cell = row[0]
            value = "" if cell.value is None else str(cell.value)
            out.append((cell.row, value))
        return out
    finally:
        wb.close()


def export_results(path: Path, sheet: str, header_row: int, col: int,
                   results: dict[int, ValidationResult]) -> Path:
    """Write Status/Error/Suggestion columns into a copy of the workbook."""
    wb = load_workbook(path)
    ws = wb[sheet]
    ncols = ws.max_column
    status_col, desc_col, sugg_col = ncols + 1, ncols + 2, ncols + 3

    bold = Font(bold=True)
    ws.cell(header_row, status_col, "Validation Status").font = bold
    ws.cell(header_row, desc_col, "Error Description").font = bold
    ws.cell(header_row, sugg_col, "Suggested Correction").font = bold

    for row_no, res in results.items():
        c_status = ws.cell(row_no, status_col, res.status.value)
        c_desc = ws.cell(row_no, desc_col, res.error_description)
        ws.cell(row_no, sugg_col, res.suggestion or "")
        email_cell = ws.cell(row_no, col)
        if res.status == Status.INVALID:
            for c in (email_cell, c_status, c_desc):
                c.fill = FILL_ERROR
                c.font = FONT_ERROR
        elif res.status == Status.VALID_WITH_WARNINGS:
            for c in (email_cell, c_status, c_desc):
                c.fill = FILL_WARNING
                c.font = FONT_WARNING

    for c, width in ((status_col, 22), (desc_col, 60), (sugg_col, 30)):
        ws.column_dimensions[get_column_letter(c)].width = width

    out_path = path.with_name(f"{path.stem}_validated.xlsx")
    wb.save(out_path)
    return out_path


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------

class EmailValidatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Email Validator — Excel")
        self.geometry("980x640")
        self.minsize(820, 520)

        self.file_path: Path | None = None
        self.location: tuple[str, int, int] | None = None  # sheet, header_row, col
        self.results: dict[int, ValidationResult] = {}
        self.email_rows: list[tuple[int, str]] = []
        self._queue: queue.Queue = queue.Queue()
        self._worker: threading.Thread | None = None

        self._build_ui()
        self.after(80, self._poll_queue)

    # ---------------- UI layout ----------------

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # --- top: file picker ---
        top = ttk.LabelFrame(self, text="1. Source Excel file")
        top.pack(fill="x", **pad)

        self.btn_browse = ttk.Button(top, text="Browse…", command=self._browse)
        self.btn_browse.pack(side="left", padx=8, pady=8)

        self.lbl_file = ttk.Label(top, text="No file selected",
                                  foreground="#666", width=70, anchor="w")
        self.lbl_file.pack(side="left", padx=4)

        # --- middle: detected location + controls ---
        ctrl = ttk.LabelFrame(self, text="2. Validate")
        ctrl.pack(fill="x", **pad)

        self.lbl_detect = ttk.Label(ctrl, text="Email column: —", anchor="w")
        self.lbl_detect.grid(row=0, column=0, columnspan=3, sticky="w",
                             padx=8, pady=(6, 2))

        self.var_dns = tk.BooleanVar(value=True)
        ttk.Checkbutton(ctrl, text="DNS domain check (needs internet)",
                        variable=self.var_dns).grid(row=1, column=0,
                                                    sticky="w", padx=8)

        self.btn_run = ttk.Button(ctrl, text="▶ Validate", command=self._run,
                                  state="disabled")
        self.btn_run.grid(row=1, column=1, padx=8, pady=6)

        self.btn_export = ttk.Button(ctrl, text="💾 Export results to Excel",
                                     command=self._export, state="disabled")
        self.btn_export.grid(row=1, column=2, padx=8, pady=6)

        self.progress = ttk.Progressbar(ctrl, mode="determinate")
        self.progress.grid(row=2, column=0, columnspan=3, sticky="ew",
                           padx=8, pady=(2, 8))
        ctrl.columnconfigure(0, weight=1)

        # --- summary ---
        summ = ttk.LabelFrame(self, text="3. Summary")
        summ.pack(fill="x", **pad)
        self.lbl_counts = ttk.Label(summ, text="—", anchor="w",
                                    font=("Segoe UI", 10, "bold"))
        self.lbl_counts.pack(side="left", padx=8, pady=8)

        ttk.Label(summ, text="Show:").pack(side="left", padx=(24, 4))
        self.var_filter = tk.StringVar(value="All")
        self.cmb_filter = ttk.Combobox(summ, textvariable=self.var_filter,
                                       values=FILTER_OPTIONS, width=10,
                                       state="readonly")
        self.cmb_filter.pack(side="left")
        self.cmb_filter.bind("<<ComboboxSelected>>", lambda _e: self._refresh_table())

        # --- results table ---
        tbl_frame = ttk.LabelFrame(self, text="4. Results")
        tbl_frame.pack(fill="both", expand=True, **pad)

        cols = ("row", "email", "status", "description", "suggestion")
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show="headings",
                                 selectmode="browse")
        self.tree.heading("row", text="Row")
        self.tree.heading("email", text="Email")
        self.tree.heading("status", text="Status")
        self.tree.heading("description", text="Error Description")
        self.tree.heading("suggestion", text="Suggestion")
        self.tree.column("row", width=60, anchor="e", stretch=False)
        self.tree.column("email", width=240)
        self.tree.column("status", width=170, stretch=False)
        self.tree.column("description", width=340)
        self.tree.column("suggestion", width=180)

        for status, (bg, fg) in ROW_TAG_COLORS.items():
            self.tree.tag_configure(status.value, background=bg, foreground=fg)

        vsb = ttk.Scrollbar(tbl_frame, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self._copy_cell)

        # --- status bar ---
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken",
                  anchor="w").pack(fill="x", side="bottom")

    # ---------------- actions ----------------

    def _browse(self):
        fname = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not fname:
            return
        self.file_path = Path(fname)
        self.lbl_file.config(text=str(self.file_path), foreground="#000")
        self.status_var.set("Detecting email column…")
        self.update_idletasks()
        try:
            loc = detect_email_sheet_column(self.file_path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Could not read file", str(exc))
            self.status_var.set("Failed to read the file.")
            return
        if not loc:
            messagebox.showwarning(
                "No emails found",
                "Could not find an email column in any sheet.\n"
                "Make sure a column header contains 'email' or that a "
                "column holds email addresses.")
            self.status_var.set("No email column detected.")
            return
        self.location = loc
        sheet, header_row, col = loc
        self.email_rows = load_emails(self.file_path, sheet, header_row, col)
        col_letter = get_column_letter(col)
        n = sum(1 for _, v in self.email_rows if v.strip())
        self.lbl_detect.config(
            text=f"Email column: sheet '{sheet}', column {col_letter} "
                 f"(header row {header_row}) — {n} non-empty value(s)")
        self.btn_run.config(state="normal")
        self.status_var.set("Ready to validate.")

    def _run(self):
        if not self.file_path or not self.location or self._worker and self._worker.is_alive():
            return
        self.results.clear()
        self.tree.delete(*self.tree.get_children())
        self.btn_run.config(state="disabled")
        self.btn_export.config(state="disabled")
        self.btn_browse.config(state="disabled")
        self.progress.config(maximum=max(1, len(self.email_rows)), value=0)
        use_dns = self.var_dns.get()
        self.status_var.set("Validating…")

        def work():
            checker = DomainChecker(enabled=use_dns)
            for i, (row_no, value) in enumerate(self.email_rows):
                res = validate_email(value, checker, check_dns=use_dns)
                self._queue.put(("row", row_no, res, i + 1))
            self._queue.put(("done",))

        self._worker = threading.Thread(target=work, daemon=True)
        self._worker.start()

    def _poll_queue(self):
        try:
            while True:
                item = self._queue.get_nowait()
                if item[0] == "row":
                    _, row_no, res, done = item
                    self.results[row_no] = res
                    self.progress.config(value=done)
                    if done % 25 == 0:
                        self.status_var.set(
                            f"Validating… {done}/{len(self.email_rows)}")
                elif item[0] == "done":
                    self._on_done()
        except queue.Empty:
            pass
        self.after(80, self._poll_queue)

    def _on_done(self):
        self.btn_run.config(state="normal")
        self.btn_browse.config(state="normal")
        self.btn_export.config(state="normal")
        counts = {s: 0 for s in Status}
        for res in self.results.values():
            counts[res.status] += 1
        total = sum(counts.values())
        self.lbl_counts.config(
            text=(f"Total: {total}   |   "
                  f"✅ Valid: {counts[Status.VALID]}   |   "
                  f"⚠️ Warnings: {counts[Status.VALID_WITH_WARNINGS]}   |   "
                  f"❌ Invalid: {counts[Status.INVALID]}"))
        self.status_var.set(
            f"Done — {counts[Status.INVALID]} invalid, "
            f"{counts[Status.VALID_WITH_WARNINGS]} with warnings.")
        self._refresh_table()

    def _refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        f = self.var_filter.get()
        for row_no in sorted(self.results):
            res = self.results[row_no]
            if f == "Valid" and res.status != Status.VALID:
                continue
            if f == "Warnings" and res.status != Status.VALID_WITH_WARNINGS:
                continue
            if f == "Invalid" and res.status != Status.INVALID:
                continue
            icon = {Status.VALID: "✅", Status.INVALID: "❌",
                    Status.VALID_WITH_WARNINGS: "⚠️"}[res.status]
            self.tree.insert(
                "", "end",
                values=(row_no, res.email, f"{icon} {res.status.value}",
                        res.error_description, res.suggestion or ""),
                tags=(res.status.value,))

    def _export(self):
        if not (self.file_path and self.location and self.results):
            return
        sheet, header_row, col = self.location
        try:
            out = export_results(self.file_path, sheet, header_row, col,
                                 self.results)
        except PermissionError:
            messagebox.showerror(
                "Cannot save",
                "The output file is open in Excel. Close it and try again.")
            return
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Export failed", str(exc))
            return
        self.status_var.set(f"Saved: {out}")
        messagebox.showinfo("Export complete", f"Results saved to:\n{out}")

    def _copy_cell(self, event):
        """Double-click a row → copy email + description to clipboard."""
        item = self.tree.identify_row(event.y)
        if not item:
            return
        vals = self.tree.item(item, "values")
        if vals:
            self.clipboard_clear()
            self.clipboard_append(f"{vals[1]}\t{vals[3]}")
            self.status_var.set(f"Copied: {vals[1]}")


def main():
    app = EmailValidatorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
