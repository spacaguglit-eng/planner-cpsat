# -*- coding: utf-8 -*-
import argparse
import datetime as dt
import csv
from win32com.client import Dispatch
from planner.events import optimize_all

def main():
    ap = argparse.ArgumentParser(description="CP-SAT Планировщик (CLI)")
    ap.add_argument("--start", required=True, help='Старт, формат "ДД.ММ.ГГГГ ЧЧ:ММ"')
    ap.add_argument("--file", default="", help="Путь к Excel файлу (если не указан — берём активную книгу)")
    ap.add_argument("--csv", default="plan.csv", help="Куда сохранить CSV с планом")
    args = ap.parse_args()

    plan_start = dt.datetime.strptime(args.start, "%d.%m.%Y %H:%M")

    excel = Dispatch("Excel.Application")
    if args.file:
        wb = excel.Workbooks.Open(args.file); wb.Activate()
    if excel.ActiveWorkbook is None:
        raise RuntimeError("Нет активной книги Excel (и файл не задан).")

    rows, line_stats, events = optimize_all(excel, plan_start)

    with open(args.csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=("Line","Pos","JobID","Name","Volume","Priority","StrictKey","Start","End"))
        w.writeheader()
        for r in rows:
            w.writerow(r)

    for line in sorted(line_stats.keys()):
        st = line_stats[line]
        print(f"[{line}] БАЗА idle: {st['base_total']:.1f} | ОПТ idle: {st['opt_total']:.1f} | ЭКОНОМИЯ: {st['saved']:.1f} ({st['saved_pct']:.1f}%)")

if __name__ == "__main__":
    main()
