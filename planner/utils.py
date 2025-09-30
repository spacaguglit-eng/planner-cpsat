# -*- coding: utf-8 -*-
import re

TIME_SCALE = 1  # минута = 1

CHANGEOVER_FALLBACK_MIN = 20.0
LAUNCH_FALLBACK_MIN     = 30.0

def sku_key_norm(s: str) -> str:
    s = (s or "").replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().upper()
    s = s.replace("Ё", "Е")
    return s

def line_header_from_name(s: str) -> str:
    s = (s or "")
    digits = "".join(ch for ch in s if ch.isdigit())
    return f"ЛИНИЯ {int(digits)}" if digits else s

def parse_mins_and_nextlaunch(v):
    mins = 0.0
    next_launch = -1.0
    if v is None:
        return mins, next_launch, False
    s = str(v).strip()
    if not s:
        return mins, next_launch, False
    if ";" in s:
        left, right = [p.strip() for p in s.split(";", 1)]
        if left.replace(".", "", 1).replace(",", "", 1).isdigit():
            mins = float(left.replace(",", "."))
        if right.replace(".", "", 1).replace(",", "", 1).isdigit():
            next_launch = float(right.replace(",", "."))
        return mins, next_launch, True
    if s.replace(".", "", 1).replace(",", "", 1).isdigit():
        mins = float(s.replace(",", "."))
        return mins, next_launch, True
    return mins, next_launch, False

def fmt_job(j: dict) -> str:
    return f"{j['Name']} {j['Volume']}"
