# -*- coding: utf-8 -*-
import datetime as dt
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from win32com.client import Dispatch  # Excel COM
from planner.events import optimize_all

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Планировщик (временная модель, OR-Tools)")
        self.geometry("1920x1080")

        top = ttk.Frame(self); top.pack(fill=tk.X, padx=8, pady=8)

        ttk.Label(top, text="Старт планирования (ДД.ММ.ГГГГ ЧЧ:ММ):").pack(side=tk.LEFT)
        self.dt_var = tk.StringVar(value=dt.datetime.now().replace(hour=8, minute=0, second=0, microsecond=0).strftime("%d.%m.%Y %H:%M"))
        ttk.Entry(top, textvariable=self.dt_var, width=18).pack(side=tk.LEFT, padx=6)

        ttk.Button(top, text="Читать активную книгу и оптимизировать", command=self.run_optimize).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Открыть файл Excel…", command=self.open_file_and_optimize).pack(side=tk.LEFT, padx=6)

        self.btn_save = ttk.Button(top, text="Сохранить таблицу в CSV…", command=self.save_csv, state=tk.DISABLED)
        self.btn_save.pack(side=tk.RIGHT)

        nb = ttk.Notebook(self); nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))

        self.tree_cols = ("Line","Pos","JobID","Name","Volume","Priority","StrictKey","Start","End")
        self.tree = ttk.Treeview(nb, columns=self.tree_cols, show="headings")
        for c in self.tree_cols:
            w = 90
            if c in ("Name","Start","End"): w = 160
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor=tk.CENTER)
        nb.add(self.tree, text="План")

        self.events_cols = ("Line","Type","Start","End","JobID","SKU","Qty","Speed","Minutes","Note")
        self.ev = ttk.Treeview(nb, columns=self.events_cols, show="headings")
        for c in self.events_cols:
            w = 90
            if c in ("SKU","Note"): w = 220
            self.ev.heading(c, text=c)
            self.ev.column(c, width=w, anchor=tk.CENTER)
        nb.add(self.ev, text="События")

        self.log = tk.Text(nb, height=10)
        nb.add(self.log, text="Лог")

        self.results = []
        self.line_stats = {}
        self.events = []

    def log_print(self, *args):
        s = " ".join(str(a) for a in args) + "\n"
        self.log.insert(tk.END, s); self.log.see(tk.END); self.update_idletasks()

    def _parse_dt(self) -> dt.datetime:
        s = self.dt_var.get().strip()
        try:
            return dt.datetime.strptime(s, "%d.%m.%Y %H:%M")
        except Exception:
            raise RuntimeError("Неверный формат даты/времени. Пример: 25.09.2025 08:00")

    def run_optimize(self):
        try:
            excel = Dispatch("Excel.Application")
            if excel.ActiveWorkbook is None:
                messagebox.showerror("Ошибка", "Открой книгу в Excel перед запуском.")
                return
            plan_start = self._parse_dt()

            self.log.delete("1.0", tk.END)
            self.log_print("Читаю Jobs/матрицы…")
            self.results, self.line_stats, self.events = optimize_all(excel, plan_start)

            self.refresh_tables()
            self.btn_save.config(state=tk.NORMAL if self.results else tk.DISABLED)

            for line in sorted(self.line_stats.keys()):
                st = self.line_stats[line]
                self.log_print(f"[{line}] БАЗА idle: {st['base_total']:.1f} мин | ОПТ idle: {st['opt_total']:.1f} мин | ЭКОНОМИЯ: {st['saved']:.1f} мин ({st['saved_pct']:.1f}%)")
                self.log_print("  Переходы (опт):")
                for d in st["opt_details"]:
                    self.log_print(f"    {d['from']} -> {d['to']} : {d['cost']:.1f} мин")
                self.log_print("  Переходы (база):")
                for d in st["base_details"]:
                    self.log_print(f"    {d['from']} -> {d['to']} : {d['cost']:.1f} мин")

            self.log_print("Готово. Строк в плане:", len(self.results), " | Событий:", len(self.events))
        except Exception as e:
            self.log_print("Ошибка:", e)
            import traceback; self.log_print(traceback.format_exc())
            messagebox.showerror("Ошибка", str(e))

    def open_file_and_optimize(self):
        path = filedialog.askopenfilename(
            title="Выбери файл Excel",
            filetypes=[("Excel files","*.xlsx;*.xlsm;*.xlsb;*.xls")]
        )
        if not path:
            return
        try:
            excel = Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(path); wb.Activate()
            self.log_print("Открыт файл:", path)
            self.run_optimize()
        except Exception as e:
            self.log_print("Ошибка:", e)
            import traceback; self.log_print(traceback.format_exc())
            messagebox.showerror("Ошибка", str(e))

    def refresh_tables(self):
        for it in self.tree.get_children(): self.tree.delete(it)
        for r in self.results:
            self.tree.insert("", tk.END, values=tuple(r.get(c, "") for c in self.tree_cols))
        for it in self.ev.get_children(): self.ev.delete(it)
        for e in self.events:
            row = {
                "Line": e["Line"], "Type": e["Type"],
                "Start": e["Start"].strftime("%d.%m %H:%M"),
                "End":   e["End"].strftime("%d.%m %H:%M"),
                "JobID": e["JobID"], "SKU": e["SKU"],
                "Qty": e["Qty"], "Speed": e["Speed"],
                "Minutes": e["Minutes"], "Note": e["Note"]
            }
            self.ev.insert("", tk.END, values=tuple(row.get(c, "") for c in self.events_cols))

    def save_csv(self):
        if not self.results:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV","*.csv")],
            title="Сохранить таблицу плана в CSV"
        )
        if not path: return
        try:
            import csv
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=list(self.tree_cols))
                w.writeheader()
                for r in self.results:
                    w.writerow(r)
            messagebox.showinfo("Сохранено", f"Файл сохранён:\n{path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    try:
        App().mainloop()
    except Exception as e:
        print("Fatal:", e)
        import sys; sys.exit(1)
