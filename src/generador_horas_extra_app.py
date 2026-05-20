#!/usr/bin/env python3
"""Interfaz gráfica simple para generar Excel de horas extra desde PDF."""

from __future__ import annotations

import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from generate_billable_hours_from_pdf import (
    build_actual_days,
    extract_records,
    make_month_rows,
    parse_date,
    parse_time,
    write_workbook,
)


class HorasExtraApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Generador de Horas Extra")
        self.geometry("720x430")
        self.minsize(680, 400)

        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.as_of = tk.StringVar(value=date.today().isoformat())
        self.no_plan = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Seleccione un PDF para comenzar.")

        self.monthly_cap = tk.StringVar(value="25")
        self.weekly_cap = tk.StringVar(value="8")
        self.daily_cap = tk.StringVar(value="3")
        self.regular_hours = tk.StringVar(value="8")
        self.late_starts_at = tk.StringVar(value="08:16:00")
        self.early_exit_before = tk.StringVar(value="16:00:00")
        self.rule_events_allowed = tk.StringVar(value="3")
        self.planned_entry = tk.StringVar(value="08:10:04")

        self._build_ui()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)

        title = ttk.Label(outer, text="Generador de Horas Extra", font=("TkDefaultFont", 18, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))

        ttk.Label(outer, text="PDF de asistencia").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(outer, textvariable=self.pdf_path).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(outer, text="Seleccionar", command=self.select_pdf).grid(row=1, column=2, sticky="ew")

        ttk.Label(outer, text="Archivo Excel de salida").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(outer, textvariable=self.output_path).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(outer, text="Guardar como", command=self.select_output).grid(row=2, column=2, sticky="ew")

        options = ttk.LabelFrame(outer, text="Opciones", padding=12)
        options.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(14, 10))
        for col in range(4):
            options.columnconfigure(col, weight=1)

        self._field(options, "Fecha de corte", self.as_of, 0, 0)
        ttk.Checkbutton(options, text="No planificar días futuros", variable=self.no_plan).grid(
            row=0, column=2, columnspan=2, sticky="w", padx=8, pady=5
        )
        self._field(options, "Tope mensual", self.monthly_cap, 1, 0)
        self._field(options, "Tope semanal", self.weekly_cap, 1, 2)
        self._field(options, "Tope diario", self.daily_cap, 2, 0)
        self._field(options, "Horas regulares", self.regular_hours, 2, 2)
        self._field(options, "Llegada tarde desde", self.late_starts_at, 3, 0)
        self._field(options, "Salida anticipada antes de", self.early_exit_before, 3, 2)
        self._field(options, "Usos de cupo", self.rule_events_allowed, 4, 0)
        self._field(options, "Entrada planificada", self.planned_entry, 4, 2)

        buttons = ttk.Frame(outer)
        buttons.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 8))
        buttons.columnconfigure(0, weight=1)
        self.generate_button = ttk.Button(buttons, text="Generar Excel", command=self.generate)
        self.generate_button.grid(row=0, column=1, sticky="e")

        status = ttk.Label(outer, textvariable=self.status, foreground="#17324D")
        status.grid(row=5, column=0, columnspan=3, sticky="w", pady=(12, 0))

    def _field(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        row: int,
        col: int,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=8, pady=5)
        ttk.Entry(parent, textvariable=variable, width=16).grid(
            row=row, column=col + 1, sticky="ew", padx=8, pady=5
        )

    def select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar PDF",
            filetypes=[("PDF", "*.pdf"), ("Todos los archivos", "*.*")],
        )
        if not path:
            return
        pdf = Path(path)
        self.pdf_path.set(str(pdf))
        if not self.output_path.get():
            self.output_path.set(str(pdf.with_name(f"{pdf.stem}_horas_extra.xlsx")))
        self.status.set("PDF seleccionado. Listo para generar.")

    def select_output(self) -> None:
        initial = self.output_path.get() or "horas_extra.xlsx"
        path = filedialog.asksaveasfilename(
            title="Guardar Excel como",
            defaultextension=".xlsx",
            initialfile=Path(initial).name,
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def generate(self) -> None:
        self.generate_button.configure(state="disabled")
        self.status.set("Generando Excel...")
        thread = threading.Thread(target=self._generate_worker, daemon=True)
        thread.start()

    def _generate_worker(self) -> None:
        try:
            output = self._generate_excel()
        except Exception as exc:  # pragma: no cover - UI error path
            self.after(0, self._show_error, exc)
            return
        self.after(0, self._show_success, output)

    def _generate_excel(self) -> Path:
        pdf = Path(self.pdf_path.get()).expanduser()
        output = Path(self.output_path.get()).expanduser()
        if not pdf.exists():
            raise ValueError("El PDF seleccionado no existe.")
        if output.suffix.lower() != ".xlsx":
            output = output.with_suffix(".xlsx")

        records = extract_records(pdf)
        year = records[0][0].year
        month = records[0][0].month
        actual_days = build_actual_days(records)
        rows = make_month_rows(
            actual_days=actual_days,
            year=year,
            month=month,
            as_of=parse_date(self.as_of.get()),
            plan_remaining=not self.no_plan.get(),
            planned_entry=parse_time(self.planned_entry.get()),
            regular_hours=float(self.regular_hours.get()),
            daily_cap=float(self.daily_cap.get()),
            weekly_cap=float(self.weekly_cap.get()),
            monthly_cap=float(self.monthly_cap.get()),
            late_starts_at=parse_time(self.late_starts_at.get()),
            early_exit_before=parse_time(self.early_exit_before.get()),
            allowed_rule_events=int(self.rule_events_allowed.get()),
        )
        write_workbook(
            rows=rows,
            output_path=output,
            monthly_cap=float(self.monthly_cap.get()),
            regular_hours=float(self.regular_hours.get()),
            weekly_cap=float(self.weekly_cap.get()),
            daily_cap=float(self.daily_cap.get()),
            late_starts_at=parse_time(self.late_starts_at.get()),
            early_exit_before=parse_time(self.early_exit_before.get()),
            allowed_rule_events=int(self.rule_events_allowed.get()),
        )
        return output

    def _show_success(self, output: Path) -> None:
        self.generate_button.configure(state="normal")
        self.status.set(f"Excel generado: {output}")
        messagebox.showinfo("Listo", f"Excel generado correctamente:\n{output}")

    def _show_error(self, exc: Exception) -> None:
        self.generate_button.configure(state="normal")
        self.status.set("No se pudo generar el Excel.")
        messagebox.showerror("Error", str(exc))


def main() -> None:
    app = HorasExtraApp()
    app.mainloop()


if __name__ == "__main__":
    main()
