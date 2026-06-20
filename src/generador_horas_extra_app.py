#!/usr/bin/env python3
"""Interfaz gráfica simple para generar Excel de horas extra desde PDF."""

from __future__ import annotations

import threading
import tkinter as tk
import sys
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

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
    RED = "#D93632"
    RED_DARK = "#B92724"
    INK = "#1E293B"
    MUTED = "#64748B"
    LINE = "#DCE3EA"
    SURFACE = "#FFFFFF"
    BACKGROUND = "#F3F6F8"
    TECH = "#17324D"

    @staticmethod
    def _resource_path(filename: str) -> Path:
        bundle_dir = getattr(sys, "_MEIPASS", None)
        if bundle_dir:
            return Path(bundle_dir) / filename
        return Path(__file__).resolve().parents[1] / filename

    def __init__(self) -> None:
        super().__init__()
        self.title("Generador de Horas Extra")
        self.geometry("940x720")
        self.minsize(840, 660)
        self.configure(background=self.BACKGROUND)

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

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("App.TFrame", background=self.BACKGROUND)
        style.configure("Surface.TFrame", background=self.SURFACE)
        style.configure(
            "Card.TLabelframe",
            background=self.SURFACE,
            bordercolor=self.LINE,
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=self.SURFACE,
            foreground=self.TECH,
            font=("TkDefaultFont", 11, "bold"),
            padding=(0, 0, 0, 6),
        )
        style.configure("Field.TLabel", background=self.SURFACE, foreground=self.INK)
        style.configure(
            "Help.TLabel",
            background=self.SURFACE,
            foreground=self.MUTED,
            font=("TkDefaultFont", 9),
        )
        style.configure(
            "Status.TLabel",
            background=self.BACKGROUND,
            foreground=self.TECH,
            font=("TkDefaultFont", 10),
        )
        style.configure(
            "Primary.TButton",
            background=self.RED,
            foreground="#FFFFFF",
            bordercolor=self.RED,
            font=("TkDefaultFont", 10, "bold"),
            padding=(20, 10),
        )
        style.map(
            "Primary.TButton",
            background=[("pressed", self.RED_DARK), ("active", "#E14642"), ("disabled", "#D9A5A3")],
            foreground=[("disabled", "#FFFFFF")],
        )
        style.configure(
            "Secondary.TButton",
            background="#F8FAFC",
            foreground=self.TECH,
            bordercolor="#BFCAD5",
            padding=(14, 8),
        )
        style.map("Secondary.TButton", background=[("active", "#EAF0F5")])
        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            foreground=self.INK,
            bordercolor="#BFCAD5",
            lightcolor="#BFCAD5",
            darkcolor="#BFCAD5",
            padding=(8, 7),
        )
        style.map("TEntry", bordercolor=[("focus", self.RED)], lightcolor=[("focus", self.RED)])
        style.configure("TCheckbutton", background=self.SURFACE, foreground=self.INK)
        style.map("TCheckbutton", background=[("active", self.SURFACE)])

    def _build_ui(self) -> None:
        accent = tk.Frame(self, background=self.RED, height=7)
        accent.grid(row=0, column=0, sticky="ew")

        header = tk.Frame(self, background=self.SURFACE, padx=28, pady=19)
        header.grid(row=1, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        tk.Label(
            header,
            text="GENERADOR DE HORAS EXTRA",
            background=self.SURFACE,
            foreground=self.INK,
            font=("TkDefaultFont", 20, "bold"),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            header,
            text="Gestión de asistencia  •  Herramienta institucional",
            background=self.SURFACE,
            foreground=self.MUTED,
            font=("TkDefaultFont", 10),
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        identity = tk.Canvas(
            header,
            width=280,
            height=62,
            background=self.TECH,
            highlightthickness=0,
            borderwidth=0,
        )
        identity.grid(row=0, column=1, rowspan=2, sticky="e")
        self._add_identity_watermark(identity)

        outer = ttk.Frame(self, padding=(28, 22, 28, 18), style="App.TFrame")
        outer.grid(row=2, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)
        outer.columnconfigure(0, weight=1)

        files = ttk.LabelFrame(outer, text="01  ARCHIVOS DEL PROCESO", padding=(18, 14), style="Card.TLabelframe")
        files.grid(row=0, column=0, sticky="ew")
        files.columnconfigure(1, weight=1)

        ttk.Label(files, text="PDF de asistencia", style="Field.TLabel").grid(row=0, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(files, textvariable=self.pdf_path).grid(row=0, column=1, sticky="ew", padx=12)
        ttk.Button(files, text="Seleccionar PDF", command=self.select_pdf, style="Secondary.TButton").grid(row=0, column=2, sticky="ew")
        ttk.Label(files, text="Reporte exportado por el reloj marcador.", style="Help.TLabel").grid(
            row=1, column=1, columnspan=2, sticky="w", padx=12, pady=(2, 10)
        )

        ttk.Label(files, text="Excel de salida", style="Field.TLabel").grid(row=2, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(files, textvariable=self.output_path).grid(row=2, column=1, sticky="ew", padx=12)
        ttk.Button(files, text="Guardar como", command=self.select_output, style="Secondary.TButton").grid(row=2, column=2, sticky="ew")
        ttk.Label(files, text="Destino del archivo .xlsx con la vista principal y el detalle.", style="Help.TLabel").grid(
            row=3, column=1, columnspan=2, sticky="w", padx=12, pady=(2, 4)
        )

        options = ttk.LabelFrame(outer, text="02  PARÁMETROS DEL CÁLCULO", padding=(18, 14), style="Card.TLabelframe")
        options.grid(row=1, column=0, sticky="ew", pady=(18, 0))
        for col in range(4):
            options.columnconfigure(col, weight=1)

        self._field(options, "Fecha de corte", self.as_of, 0, 0, "Ejemplo: 2026-05-31")
        ttk.Checkbutton(options, text="No planificar días futuros", variable=self.no_plan).grid(
            row=0, column=2, columnspan=2, sticky="w", padx=8, pady=5
        )
        ttk.Label(options, text="Desactiva la proyección del resto del mes.", style="Help.TLabel").grid(
            row=1, column=2, columnspan=2, sticky="w", padx=8, pady=(0, 5)
        )
        self._field(options, "Tope mensual", self.monthly_cap, 1, 0, "Máximo mensual en horas. Ejemplo: 25")
        self._field(options, "Hora oficial de entrada", self.official_entry, 1, 2, "Ejemplo: 08:00:00")
        self._field(options, "Llegada tarde desde", self.late_starts_at, 2, 0, "Primer minuto tardío. Ejemplo: 08:16:00")
        self._field(options, "Salida anticipada antes de", self.early_exit_before, 2, 2, "Ejemplo: 16:00:00")
        self._field(options, "Entrada planificada", self.planned_entry, 3, 0, "Hora usada para proyectar días futuros.")
        self._field(
            options,
            "Fechas de lluvia y tolerancia",
            self.rain_dates,
            3,
            2,
            "Cargue minutos adicionales. Ej.: 04/05/2026:30, 12/05/2026:45",
        )

        footer = ttk.Frame(outer, style="App.TFrame")
        footer.grid(row=2, column=0, sticky="ew", pady=(18, 0))
        footer.columnconfigure(0, weight=1)
        status = ttk.Label(footer, textvariable=self.status, style="Status.TLabel")
        status.grid(row=0, column=0, sticky="w")
        self.generate_button = ttk.Button(
            footer, text="GENERAR EXCEL", command=self.generate, style="Primary.TButton"
        )
        self.generate_button.grid(row=0, column=1, sticky="e")

    def _add_identity_watermark(self, canvas: tk.Canvas) -> None:
        logo_path = self._resource_path("logo_conacom_2023_.png")
        if logo_path.exists():
            logo = Image.open(logo_path).convert("RGB")
            logo.thumbnail((260, 50), Image.Resampling.LANCZOS)
            luminance = logo.convert("L")
            alpha = luminance.point(lambda value: int((255 - value) * 0.16))
            watermark = Image.new("RGBA", logo.size, "#FFFFFF")
            watermark.putalpha(alpha)
            self.identity_logo = ImageTk.PhotoImage(watermark)
            canvas.create_image(140, 31, image=self.identity_logo)
        canvas.create_text(
            140,
            31,
            text="CONACOM  |  CAPITAL HUMANO",
            fill="#FFFFFF",
            font=("TkDefaultFont", 9, "bold"),
        )

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
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=grid_row, column=col, sticky="w", padx=8, pady=(5, 0))
        ttk.Entry(parent, textvariable=variable, width=16, state=state).grid(
            row=grid_row, column=col + 1, sticky="ew", padx=8, pady=(5, 0)
        )
        ttk.Label(parent, text=help_text, style="Help.TLabel", wraplength=270).grid(
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
