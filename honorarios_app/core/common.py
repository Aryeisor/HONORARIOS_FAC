import re


def norm(s) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip().upper()


def normalize_datetime_key(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    s = s.replace(".0", "")
    return s


def parse_percent(x):
    if x is None or str(x).strip() == "":
        return None

    if isinstance(x, (int, float)):
        v = float(x)

        # Mantener valores como:
        # 0.50, 0.75, 1.00, 1.05, 1.75
        if 0 <= v <= 2:
            return v

        # Convertir valores como 50, 75, 90, 105, 175
        return v / 100.0

    s = str(x).strip().replace(",", ".").replace("%", "")
    try:
        v = float(s)
    except ValueError:
        return None

    if 0 <= v <= 2:
        return v

    return v / 100.0


def to_number(x) -> float:
    if x is None or x == "":
        return 0.0

    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def normalize_fac(value) -> str:
    s = norm(value)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def find_header_row(ws, required_headers):
    max_scan = min(ws.max_row, 200)

    for r in range(1, max_scan + 1):
        col_map = {}
        for c in range(1, ws.max_column + 1):
            v = norm(ws.cell(r, c).value)
            if v:
                col_map[v] = c

        if all(any(req in h for h in col_map.keys()) for req in required_headers):
            return r, col_map

    return None, {}


def get_col(col_map, *candidates):
    keys = list(col_map.keys())

    for cand in candidates:
        cand_n = norm(cand)

        if cand_n in col_map:
            return col_map[cand_n]

        for k in keys:
            if cand_n in k:
                return col_map[k]

    return None