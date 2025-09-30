# -*- coding: utf-8 -*-
def read_jobs_from_active_excel(excel_app):
    wb = excel_app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("Нет активной книги Excel.")

    try:
        ws = wb.Worksheets("JOBS")
    except Exception:
        raise RuntimeError("Лист 'JOBS' не найден.")

    lo = None
    for L in ws.ListObjects:
        if str(L.Name).lower() == "jobstable":
            lo = L
            break
    if lo is None or lo.DataBodyRange is None:
        raise RuntimeError("Таблица 'JobsTable' на листе 'JOBS' не найдена.")

    headers = [str(h.Value or "").strip() for h in lo.HeaderRowRange]

    def find_col(*names):
        for idx, name in enumerate(headers, start=1):
            nm = str(name or "").upper().replace("\u00A0", " ").strip()
            for want in names:
                if nm == str(want or "").upper():
                    return idx
        return None

    col_job   = find_col("JOBID")
    col_line  = find_col("LINE")
    col_name  = find_col("NAME")
    col_vol   = find_col("VOLUME")
    col_qty   = find_col("QUANTITY")
    col_speed = find_col("SPEED")
    col_prio  = find_col("PRIORITY")
    col_strict= find_col("СТРОГИЙ ПОРЯДОК", "STRICTKEY")

    need = [("JobID", col_job), ("Line", col_line), ("Name", col_name),
            ("Volume", col_vol), ("Quantity", col_qty), ("Speed", col_speed),
            ("Priority", col_prio)]
    missing = [nm for nm, c in need if not c]
    if missing:
        raise RuntimeError("Нет колонок в JobsTable: " + ", ".join(missing))

    rows = lo.DataBodyRange.Rows.Count
    jobs = []
    for r in range(1, rows + 1):
        def val(c):
            return lo.DataBodyRange.Cells(r, c).Value if c else ""
        speed = float(val(col_speed) or 0)
        qty = float(val(col_qty) or 0)
        if speed <= 0 or qty <= 0:
            continue
        job = {
            "JobID":    str(val(col_job) or "").strip(),
            "Line":     str(val(col_line) or "").strip(),
            "Name":     str(val(col_name) or "").strip(),
            "Volume":   str(val(col_vol) or "").strip(),
            "Quantity": qty,
            "Speed":    speed,
            "Priority": int(float(val(col_prio) or 0)),
            "StrictKey": str(val(col_strict) or "").strip() if col_strict else "",
            "_row": r,
        }
        if job["JobID"] and job["Line"]:
            jobs.append(job)
    return jobs
