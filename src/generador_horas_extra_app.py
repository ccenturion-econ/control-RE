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
    parse_rain_dates,
    parse_time,
    write_workbook,
)


class HorasExtraApp(tk.Tk):
    MONTHLY_CAP = 25.0
    WEEKLY_CAP = 8.0
    DAILY_CAP = 3.0
    REGULAR_HOURS = 8.0
    RULE_EVENTS_ALLOWED = 3

    def __init__(self) -> None:
        super().__init__()
        self.title("Generador de Horas Extra")
        self.geometry("860x650")
        self.minsize(780, 600)

        self.pdf_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.as_of = tk.StringVar(value=date.today().isoformat())
        self.no_plan = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Seleccione un PDF para comenzar.")

        self.monthly_cap = tk.StringVar(value="25")
        self.official_entry = tk.StringVar(value="08:00:00")
        self.late_starts_at = tk.StringVar(value="08:16:00")
        self.early_exit_before = tk.StringVar(value="16:00:00")
        self.planned_entry = tk.StringVar(value="08:10:04")
        self.rain_dates = tk.StringVar()

        self._build_ui()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=18)
        outer.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)

        title = ttk.Label(outer, text="Generador de Horas Extra", font=("TkDefaultFont", 18, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 18))

        ttk.Label(outer, text="PDF de asistencia").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(outer, textvariable=self.pdf_path).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(outer, text="Seleccionar", command=self.select_pdf).grid(row=1, column=2, sticky="ew")
        ttk.Label(outer, text="Seleccione el reporte PDF exportado por el reloj marcador.", foreground="#607080").grid(
            row=2, column=1, columnspan=2, sticky="w", padx=8, pady=(0, 6)
        )

        ttk.Label(outer, text="Archivo Excel de salida").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(outer, textvariable=self.output_path).grid(row=3, column=1, sticky="ew", padx=8)
        ttk.Button(outer, text="Guardar como", command=self.select_output).grid(row=3, column=2, sticky="ew")
        ttk.Label(outer, text="Indique dónde guardar el archivo .xlsx generado.", foreground="#607080").grid(
            row=4, column=1, columnspan=2, sticky="w", padx=8, pady=(0, 6)
        )

        options = ttk.LabelFrame(outer, text="Opciones", padding=12)
        options.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(14, 10))
        for col in range(4):
            options.columnconfigure(col, weight=1)

        self._field(options, "Fecha de corte", self.as_of, 0, 0, "Ejemplo: 2026-05-31")
        ttk.Checkbutton(options, text="No planificar días futuros", variable=self.no_plan).grid(
            row=0, column=2, columnspan=2, sticky="w", padx=8, pady=5
        )
        ttk.Label(options, text="Desactive la proyección del resto del mes.", foreground="#607080").grid(
            row=1, column=2, columnspan=2, sticky="w", padx=8, pady=(0, 5)
        )
        self._field(options, "Tope mensual", self.monthly_cap, 1, 0, "Máximo mensual en horas. Ejemplo: 25")
        self._field(options, "Hora oficial de entrada", self.official_entry, 1, 2, "Ejemplo: 08:00:00")
        self._field(options, "Llegada tarde desde", self.late_starts_at, 2, 0, "Primer minuto tardío. Ejemplo: 08:16:00")
        self._field(options, "Salida anticipada antes de", self.early_exit_before, 2, 2, "Ejemplo: 16:00:00")
        self._field(options, "Entrada planificada", self.planned_entry, 3, 0, "Hora usada para proyectar días futuros.")
        self._field(options, "Fechas de lluvia", self.rain_dates, 3, 2, "Día/mes/año, separadas por comas. Ej.: 04/05/2026, 12/05/2026")

        buttons = ttk.Frame(outer)
        buttons.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(14, 8))
        buttons.columnconfigure(0, weight=1)
        self.generate_button = ttk.Button(buttons, text="Generar Excel", command=self.generate)
        self.generate_button.grid(row=0, column=1, sticky="e")

        status = ttk.Label(outer, textvariable=self.status, foreground="#17324D")
        status.grid(row=7, column=0, columnspan=3, sticky="w", pady=(12, 0))

    def _field(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        row: int,
        col: int,
        help_text: str,
        state: str = "normal",
    ) -> None:
        grid_row = row * 2
        ttk.Label(parent, text=label).grid(row=grid_row, column=col, sticky="w", padx=8, pady=(5, 0))
        ttk.Entry(parent, textvariable=variable, width=16, state=state).grid(
            row=grid_row, column=col + 1, sticky="ew", padx=8, pady=(5, 0)
        )
        ttk.Label(parent, text=help_text, foreground="#607080", wraplength=250).grid(
            row=grid_row + 1, column=col, columnspan=2, sticky="w", padx=8, pady=(0, 5)
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
        rain_dates = parse_rain_dates(self.rain_dates.get())
        rows = make_month_rows(
            actual_days=actual_days,
            year=year,
            month=month,
            as_of=parse_date(self.as_of.get()),
            plan_remaining=not self.no_plan.get(),
            planned_entry=parse_time(self.planned_entry.get()),
            official_entry=parse_time(self.official_entry.get()),
            regular_hours=self.REGULAR_HOURS,
            daily_cap=self.DAILY_CAP,
            weekly_cap=self.WEEKLY_CAP,
            monthly_cap=float(self.monthly_cap.get()),
            late_starts_at=parse_time(self.late_starts_at.get()),
            early_exit_before=parse_time(self.early_exit_before.get()),
            allowed_rule_events=self.RULE_EVENTS_ALLOWED,
            rain_dates=rain_dates,
        )
        write_workbook(
            rows=rows,
            output_path=output,
            monthly_cap=float(self.monthly_cap.get()),
            regular_hours=self.REGULAR_HOURS,
            weekly_cap=self.WEEKLY_CAP,
            daily_cap=self.DAILY_CAP,
            official_entry=parse_time(self.official_entry.get()),
            late_starts_at=parse_time(self.late_starts_at.get()),
            early_exit_before=parse_time(self.early_exit_before.get()),
            allowed_rule_events=self.RULE_EVENTS_ALLOWED,
            rain_dates=rain_dates,
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
