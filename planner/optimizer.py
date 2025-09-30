# -*- coding: utf-8 -*-
import math
from ortools.sat.python import cp_model
from .utils import (
    sku_key_norm, TIME_SCALE, CHANGEOVER_FALLBACK_MIN, LAUNCH_FALLBACK_MIN, fmt_job
)

def _proc_min(job: dict) -> int:
    qty = float(job.get("Quantity", 0) or 0.0)
    spd = float(job.get("Speed", 0) or 0.0)
    dur = 0.0 if spd <= 0 else qty / spd * 60.0
    return int(math.ceil(dur / TIME_SCALE))

def _step_cost_min(a: dict, b: dict, trans_for_line: dict) -> int:
    from_key = sku_key_norm(f"{a['Name']} {a['Volume']}")
    to_key   = sku_key_norm(f"{b['Name']} {b['Volume']}")
    if from_key == to_key:
        return 0
    rec = trans_for_line.get(f"{from_key}>>{to_key}")
    if rec:
        setup, next_launch, _ = rec
        if setup <= 0.5:
            return 0
        cost = setup + (next_launch if next_launch >= 0 else LAUNCH_FALLBACK_MIN)
    else:
        cost = CHANGEOVER_FALLBACK_MIN + LAUNCH_FALLBACK_MIN
    return int(math.ceil(cost / TIME_SCALE))

def build_line_schedule_cp(
    line_name: str,
    jobs_for_line: list[dict],
    trans_for_line: dict,
    start_launch_min: float,
    solver_time_limit_sec: float = 10.0,
    log_fn = print,
):
    n = len(jobs_for_line)
    if n == 0:
        return [], {}, 0

    durations = [_proc_min(j) for j in jobs_for_line]
    idx_by_id = {j["JobID"]: i for i, j in enumerate(jobs_for_line)}

    setup = [[0]*n for _ in range(n)]
    max_setup = 0
    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            s = _step_cost_min(jobs_for_line[i], jobs_for_line[j], trans_for_line)
            setup[i][j] = s
            if s > max_setup:
                max_setup = s

    sum_dur = sum(durations)
    start_launch = int(math.ceil(start_launch_min / TIME_SCALE))
    base_horizon = sum_dur + n * max(1, max_setup) + start_launch
    H = max(base_horizon * 2, 10_000)

    def solve_with_horizon(h: int):
        model = cp_model.CpModel()
        starts = [model.NewIntVar(0, h, f"s_{i}") for i in range(n)]
        ends   = [model.NewIntVar(0, h, f"e_{i}") for i in range(n)]
        for i in range(n):
            model.Add(ends[i] == starts[i] + durations[i])

        for i in range(n):
            for j in range(i + 1, n):
                o = model.NewBoolVar(f"o_{i}_{j}")
                model.Add(ends[i] + setup[i][j] <= starts[j]).OnlyEnforceIf(o)
                model.Add(ends[j] + setup[j][i] <= starts[i]).OnlyEnforceIf(o.Not())

        from collections import defaultdict
        buckets = defaultdict(list)
        for j in jobs_for_line:
            key = str(j.get("StrictKey", "") or "").strip()
            if key:
                buckets[key].append(j)
        for key, lst in buckets.items():
            lst_sorted = sorted(lst, key=lambda x: x["_row"])
            for a, b in zip(lst_sorted, lst_sorted[1:]):
                ia = idx_by_id[a["JobID"]]; ib = idx_by_id[b["JobID"]]
                model.Add(ends[ia] + setup[ia][ib] <= starts[ib])

        for i in range(n):
            for j in range(n):
                if jobs_for_line[i]["Priority"] < jobs_for_line[j]["Priority"]:
                    model.Add(ends[i] + setup[i][j] <= starts[j])

        makespan = model.NewIntVar(0, h, "makespan")
        model.AddMaxEquality(makespan, ends)
        model.Minimize(makespan)

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = solver_time_limit_sec
        solver.parameters.num_search_workers = 8
        status = solver.Solve(model)
        return status, solver, starts, ends

    if log_fn:
        log_fn(f"[{line_name}] sum_dur={sum_dur} мин; max_setup={max_setup} мин; "
               f"start_launch={start_launch} мин; horizon={H}")

    status, solver, starts, ends = solve_with_horizon(H)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        if log_fn: log_fn("Нет решения на H=", H, " — пробуем x10…")
        status, solver, starts, ends = solve_with_horizon(H * 10)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("CP-SAT не нашёл допустимое решение.")

    start_vals = [solver.Value(s) for s in starts]
    order_idx  = sorted(range(n), key=lambda i: start_vals[i])
    order      = [jobs_for_line[i] for i in order_idx]
    times      = {jobs_for_line[i]["JobID"]: (solver.Value(starts[i]), solver.Value(ends[i])) for i in range(n)}
    makespan   = max(solver.Value(e) for e in ends)
    return order, times, makespan

def analyze_sequence_cost(line_name: str, seq: list[dict], trans_for_line: dict):
    total = 0.0
    details = []
    if len(seq) <= 1:
        return 0.0, details
    for i in range(len(seq) - 1):
        a, b = seq[i], seq[i + 1]
        c = _step_cost_min(a, b, trans_for_line)
        total += c
        details.append({"from": fmt_job(a), "to": fmt_job(b), "cost": float(round(c, 3))})
    return float(round(total, 3)), details
