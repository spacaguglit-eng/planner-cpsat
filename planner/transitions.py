# -*- coding: utf-8 -*-
from .utils import (
    sku_key_norm, line_header_from_name, parse_mins_and_nextlaunch,
    CHANGEOVER_FALLBACK_MIN, LAUNCH_FALLBACK_MIN
)

def read_stdstops_dict(excel_app):
    std = {}
    wb = excel_app.ActiveWorkbook
    try:
        ws = wb.Worksheets("Таблица_Нормативов")
    except Exception:
        return std

    lo = None
    for L in ws.ListObjects:
        if str(L.Name).lower() == "stdstops":
            lo = L
            break
    if lo is None or lo.DataBodyRange is None:
        return std

    hdr_vals = [str(c.Value or "") for c in lo.HeaderRowRange]
    col_event = None
    for idx, name in enumerate(hdr_vals, start=1):
        nm = str(name or "").strip().lower()
        if nm in ("event", "событие"):
            col_event = idx
            break
    if not col_event:
        return std

    body = lo.DataBodyRange.Value
    n_rows = len(body)
    n_cols = len(body[0]) if n_rows else 0

    for col in range(1, n_cols + 1):
        if col == col_event:
            continue
        raw_hdr = hdr_vals[col - 1]
        if not raw_hdr:
            continue
        line_hdr = line_header_from_name(raw_hdr)
        if not line_hdr:
            continue
        line_dict = std.setdefault(line_hdr, {})
        for r in range(n_rows):
            event_key = str(body[r][col_event - 1] or "").strip()
            if not event_key:
                continue
            cell_val = body[r][col - 1]
            mins, next_launch, ok = parse_mins_and_nextlaunch(cell_val)
            if ok:
                line_dict[event_key] = (float(mins), float(next_launch))
            elif isinstance(cell_val, (int, float)):
                line_dict[event_key] = (float(cell_val), -1.0)
    return std

def read_transition_matrix_from_active_excel(excel_app):
    wb = excel_app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("Нет активной книги Excel.")
    try:
        ws = wb.Worksheets("Карта_Переходов")
    except Exception:
        raise RuntimeError("Лист 'Карта_Переходов' не найден.")

    std = read_stdstops_dict(excel_app)

    used = ws.UsedRange
    vals = used.Value
    if not vals:
        return {"transitions": {}, "start_launch": {}}

    n_rows = len(vals)
    n_cols = len(vals[0]) if n_rows else 0

    trans = {}
    r = 0
    first_data_col = 1

    while r < n_rows:
        cell_txt = str(vals[r][0] or "")
        if cell_txt and "Линия" in cell_txt:
            line_name = cell_txt.replace("Линия:", "").strip() or cell_txt.strip()
            head_row = r + 1

            c = first_data_col
            while c < n_cols and str(vals[head_row][c] or "").strip():
                c += 1
            end_col = max(c - 1, first_data_col)

            r2 = head_row + 1
            while r2 < n_rows:
                v = str(vals[r2][0] or "")
                if (not v) or ("Линия" in v):
                    break
                r2 += 1
            end_row = max(r2 - 1, head_row + 1)

            line_dict = trans.setdefault(line_name, {})
            line_hdr = line_header_from_name(line_name)
            to_headers = [sku_key_norm(str(vals[head_row][cc] or "")) for cc in range(first_data_col, end_col + 1)]

            for rr in range(head_row + 1, end_row + 1):
                from_key = sku_key_norm(str(vals[rr][0] or ""))
                if not from_key:
                    continue
                for idx_cc, cc in enumerate(range(first_data_col, end_col + 1)):
                    to_key = to_headers[idx_cc]
                    if not to_key:
                        continue
                    co_cell = vals[rr][cc]
                    mins = CHANGEOVER_FALLBACK_MIN
                    next_launch = -1.0
                    note = "Найден в линии, но пусто"
                    if co_cell is not None and str(co_cell).strip():
                        event_key = str(co_cell).strip()
                        if (line_hdr in std) and (event_key in std[line_hdr]):
                            mins, next_launch = std[line_hdr][event_key]
                            note = f"Ключ: {event_key}"
                        else:
                            note = f"Ключ '{event_key}' не найден в StdStops"
                    line_dict[f"{from_key}>>{to_key}"] = (float(mins), float(next_launch), note)
            r = end_row + 1
        else:
            r += 1

    start_launch_by_line = {}
    for line in trans.keys():
        hdr = line_header_from_name(line)
        mins, nextl = LAUNCH_FALLBACK_MIN, -1.0
        if hdr in std and "Запуск линии" in std[hdr]:
            mins, nextl = std[hdr]["Запуск линии"]
        start_launch_by_line[line] = (mins, nextl)

    return {"transitions": trans, "start_launch": start_launch_by_line}
