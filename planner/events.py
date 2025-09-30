# -*- coding: utf-8 -*-
import datetime as dt
from .utils import sku_key_norm, LAUNCH_FALLBACK_MIN, fmt_job
from .optimizer import build_line_schedule_cp, analyze_sequence_cost
from .excel_io import read_jobs_from_active_excel
from .transitions import read_transition_matrix_from_active_excel

def build_events_for_line(line_name: str,
                          order: list[dict],
                          times: dict,
                          trans_for_line: dict,
                          plan_start_dt: dt.datetime,
                          start_launch_min: float) -> list[dict]:
    events = []
    t0 = plan_start_dt + dt.timedelta(minutes=start_launch_min)

    if start_launch_min > 0.5:
        events.append({
            "Line": line_name, "Type": "Запуск",
            "Start": plan_start_dt, "End": t0,
            "JobID": "", "SKU": "", "Qty": "", "Speed": "",
            "Minutes": round(start_launch_min, 1), "Note": "Старт линии"
        })

    seq = sorted(order, key=lambda j: times[j["JobID"]][0])
    prev = None
    for j in seq:
        s_rel, e_rel = times[j["JobID"]]
        s = t0 + dt.timedelta(minutes=s_rel)
        e = t0 + dt.timedelta(minutes=e_rel)
        if prev is not None:
            from_key = sku_key_norm(f"{prev['Name']} {prev['Volume']}")
            to_key   = sku_key_norm(f"{j['Name']} {j['Volume']}")
            rec = trans_for_line.get(f"{from_key}>>{to_key}")
            if rec:
                setup, next_launch, note = rec
                if setup > 0.5:
                    extra = next_launch if next_launch >= 0 else LAUNCH_FALLBACK_MIN
                    tr_min = setup + extra
                    tr_beg = (t0 + dt.timedelta(minutes=times[prev["JobID"]][1]))
                    tr_end = tr_beg + dt.timedelta(minutes=tr_min)
                    events.append({
                        "Line": line_name, "Type": "Переход",
                        "Start": tr_beg, "End": tr_end,
                        "JobID": j["JobID"], "SKU": fmt_job(j), "Qty": "",
                        "Speed": "", "Minutes": round(tr_min, 1), "Note": note
                    })
        qty = j["Quantity"]; spd = j["Speed"]
        dur_min = (qty / spd) * 60.0
        events.append({
            "Line": line_name, "Type": "Производство",
            "Start": s, "End": e, "JobID": j["JobID"], "SKU": fmt_job(j),
            "Qty": round(qty, 0), "Speed": round(spd, 2),
            "Minutes": round(dur_min, 1), "Note": ""
        })
        prev = j

    return events

def optimize_all(excel_app, plan_start_dt: dt.datetime):
    jobs = read_jobs_from_active_excel(excel_app)
    tdata = read_transition_matrix_from_active_excel(excel_app)
    trans_all = tdata["transitions"]
    start_launch_all = tdata["start_launch"]

    from collections import defaultdict
    by_line = defaultdict(list)
    for j in jobs:
        by_line[j["Line"]].append(j)

    table_rows = []
    line_stats = {}
    all_events = []

    for line, jlist in by_line.items():
        trans_for_line = trans_all.get(line, {})
        start_launch_min = float(start_launch_all.get(line, (LAUNCH_FALLBACK_MIN, -1.0))[0])

        base_seq = sorted(jlist, key=lambda x: x["_row"])
        base_total, base_details = analyze_sequence_cost(line, base_seq, trans_for_line)

        order, times, obj = build_line_schedule_cp(line, jlist, trans_for_line, start_launch_min)

        events = build_events_for_line(line, order, times, trans_for_line, plan_start_dt, start_launch_min)
        all_events.extend(events)

        sum_prod = sum((j["Quantity"] / j["Speed"]) * 60.0 for j in jlist)
        opt_total = (obj + start_launch_min)
        idle_opt = opt_total - sum_prod
        idle_base = base_total + start_launch_min
        saved = idle_base - idle_opt
        saved_pct = (saved / idle_base * 100.0) if idle_base > 0 else 0.0

        line_stats[line] = {
            "base_total": round(idle_base, 1),
            "opt_total":  round(idle_opt, 1),
            "saved":      round(saved, 1),
            "saved_pct":  round(saved_pct, 1),
            "base_details": base_details,
            "opt_details":  analyze_sequence_cost(line, order, trans_for_line)[1],
        }

        ordered = sorted(order, key=lambda j: times[j["JobID"]][0])
        for rank, j in enumerate(ordered, start=1):
            s_rel, e_rel = times[j["JobID"]]
            table_rows.append({
                "Line": line,
                "Pos": rank,
                "JobID": j["JobID"],
                "Name": j["Name"],
                "Volume": j["Volume"],
                "Priority": j["Priority"],
                "StrictKey": j["StrictKey"],
                "Start": (plan_start_dt + dt.timedelta(minutes=start_launch_min + s_rel)).strftime("%d.%m %H:%M"),
                "End":   (plan_start_dt + dt.timedelta(minutes=start_launch_min + e_rel)).strftime("%d.%m %H:%M"),
            })

    table_rows.sort(key=lambda r: (r["Line"], r["Pos"]))
    all_events.sort(key=lambda e: (e["Line"], e["Start"]))
    return table_rows, line_stats, all_events
