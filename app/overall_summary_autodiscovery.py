# -*- coding: utf-8 -*-
"""
app/overall_summary_autodiscovery.py

Автогенерация формул для листа "Общая сводка все филиалы" на основе discovery
блоков в листах "Сводка ...".

Overall (C:L):
C  Бюджет $
D  Лидов
E  Пришло
F  Записей (все)
G  Записи ИИ
H  Записи ОП
I  Цена упавшего лида     = C/D
J  Цена дошедшего лида     = C/E
K  Возврат средств / Оплата за первичку $
L  Продажи $

КОНФИГ-ЯЧЕЙКИ (строка 1 листа Overall):
  $B$1  режим периода ("Весь период" / "Определенная дата" / "Диапазон дат")
  $C$1  дата ОТ
  $D$1  дата ДО
  $E$1  курс USD->AED (для Дубай / CocoAge), напр. 3.6725
  $F$1  курс USD->KZT (для филиалов KZ),       напр. 486.62
  $G$1  переключатель валюты ("USD" / "KZT")
  $H$1  курс USD->UZS (для Ташкента),          напр. 12024.11

ВАЛЮТА (важно):
  K (платники / возврат):
    - kz       -> делим на $F$1 (KZT)   ТОЛЬКО когда $G$1="USD"
    - tashkent -> делим на $H$1 (UZS)   ТОЛЬКО когда $G$1="USD"
    - dubai    -> делим на $E$1 (AED)   всегда
  L (продажи):
    - kz       -> делим на $F$1 (KZT)   ТОЛЬКО когда $G$1="USD"
    - tashkent -> НЕ делим (продажи уже в USD)
    - dubai    -> делим на $E$1 (AED)   всегда

СЕКЦИИ листа Overall (метки в B повторяются между секциями!):
  - "Филиал ..."  таблица    -> своя строка "Итог:"
  - "Франшиза ..." таблица    -> своя строка "Итог:"
  - "KPI" блок                -> НЕ ТРОГАЕМ (там свои формулы). Обход
                                 останавливается на строке с меткой "KPI".

Кол-во строк дат:
  - Июнь / 30 дней -> --work_rows "4:33"
  - Февраль / 28   -> --work_rows "4:31"
"""

from __future__ import annotations

import argparse
import json
import os
import re
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from google.oauth2 import service_account
from googleapiclient.discovery import build


# =========================
# ENV LOAD
# =========================
def _try_load_env() -> None:
    try:
        from dotenv import load_dotenv, find_dotenv  # type: ignore
        p = find_dotenv()
        if p:
            load_dotenv(p, override=False)
            return
    except Exception:
        pass
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        cand = os.path.join(os.path.dirname(here), ".env")
        if os.path.exists(cand):
            with open(cand, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    k = k.strip()
                    v = v.strip().strip('"').strip("'")
                    os.environ.setdefault(k, v)
    except Exception:
        pass


# =========================
# A1 utils
# =========================
def col_to_index(col: str) -> int:
    col = col.upper()
    idx = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Bad col: {col}")
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def index_to_col(idx: int) -> str:
    if idx <= 0:
        raise ValueError("idx must be >= 1")
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def parse_scan(scan: str) -> Tuple[str, str]:
    scan = scan.strip()
    if ":" not in scan:
        raise ValueError("scan must be like B:ET")
    a, b = scan.split(":", 1)
    return a.strip().upper(), b.strip().upper()


def parse_work_rows(work_rows: str) -> Tuple[int, int]:
    work_rows = work_rows.strip()
    if ":" not in work_rows:
        raise ValueError("work_rows must be like 4:33")
    a, b = work_rows.split(":", 1)
    return int(a), int(b)


def quote_sheet(sheet_name: str) -> str:
    return f"'{sheet_name}'"


def _parse_row_num(a1_part: str) -> int:
    m = re.findall(r"\d+", a1_part)
    if not m:
        raise ValueError(f"Can't parse row from: {a1_part}")
    return int(m[0])


def _parse_a1_row_range(a1: str) -> Tuple[int, int]:
    a1 = a1.replace("$", "").strip()
    if ":" not in a1:
        r = _parse_row_num(a1)
        return r, r
    a, b = a1.split(":", 1)
    return _parse_row_num(a), _parse_row_num(b)


# =========================
# Google Sheets API
# =========================
_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _get_service():
    _try_load_env()
    sa_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_path or not os.path.exists(sa_path):
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON is not set or file does not exist.")
    creds = service_account.Credentials.from_service_account_file(sa_path, scopes=_SCOPES)
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _get_sheets_properties(service, spreadsheet_id: str) -> Dict[str, Dict[str, Any]]:
    resp = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))"
    ).execute()
    out: Dict[str, Dict[str, Any]] = {}
    for sh in resp.get("sheets", []):
        p = sh.get("properties", {})
        title = p.get("title")
        gp = p.get("gridProperties", {}) or {}
        if title:
            out[title] = {
                "sheetId": p.get("sheetId"),
                "rowCount": gp.get("rowCount"),
                "colCount": gp.get("columnCount"),
            }
    return out


def batch_get_formatted(service, spreadsheet_id: str, ranges: List[str]) -> Dict[str, List[List[Any]]]:
    resp = service.spreadsheets().values().batchGet(
        spreadsheetId=spreadsheet_id,
        ranges=ranges,
        valueRenderOption="FORMATTED_VALUE"
    ).execute()
    out: Dict[str, List[List[Any]]] = {}
    for vr in resp.get("valueRanges", []):
        r = vr.get("range")
        vals = vr.get("values", [])
        if r:
            out[r] = vals
    return out


def batch_update_user_entered(service, spreadsheet_id: str, updates: List[Tuple[str, List[List[str]]]]) -> Dict[str, Any]:
    body = {
        "valueInputOption": "USER_ENTERED",
        "data": [{"range": r, "majorDimension": "ROWS", "values": v} for (r, v) in updates],
    }
    return service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()


# =========================
# Normalizers
# =========================
def _s(x: Any) -> str:
    return "" if x is None else str(x)


def _norm_base(s: str) -> str:
    s = s.replace(" ", " ")
    s = s.replace("\n", " ")
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("ё", "е")
    s = s.replace("фэшн", "фешн")   # фэшн/фешн одинаково
    s = s.replace("дубаи", "дубай")  # дубаи/дубай одинаково
    return s


def _norm_label_key(x: Any) -> str:
    return _norm_base(_s(x))


def _norm_h(x: Any) -> str:
    return _norm_base(_s(x))


def _clean_label(x: Any) -> str:
    return _s(x).replace(" ", " ").strip()


def _is_instagram_label(label: str) -> bool:
    n = _norm_label_key(label)
    return ("инстаграм" in n) or ("instagram" in n)


# служебные строки в overall (не филиалы): заголовки секций, маркеры режимов
IGNORE_LABEL_KEYS = {
    "весь период",
    "определенная дата",
    "определённая дата",
    "диапазон дат",
    "интервал дат",
    "филиал",     # заголовок столбца таблицы "Филиал"
    "франшиза",   # заголовок столбца таблицы "Франшиза"
}

# метка, на которой прекращаем обход (KPI-блок ниже трогать НЕЛЬЗЯ)
STOP_LABEL_KEYS = {"kpi"}


# =========================
# Discovery tokens
# =========================
DATE_TOKENS = {"дата", "date"}
LEADS_TOKENS = {"лидов", "лиды", "leads"}
BUDGET_TOKENS = {"бюджет", "budget"}
ARRIVED_TOKENS = {"демо", "приход", "пришло", "arrived", "demo"}
BOOK_ALL_TOKENS = {"записей", "записи", "записей (все)", "записи (все)", "bookings", "bookings all"}
# ВАЖНО: убрали голое "ai" чтобы не ловить подстроки где попало
BOOK_AI_TOKENS = {"записи ии", "записей ии", "ai bookings", "bookings ai"}
BOOK_OP_TOKENS = {"записи оп", "запись оп", "записи операторов", "запись операторов", "operator bookings", "op bookings"}

PAYERS_TOKENS = {
    "платники",
    "оплата",
    "оплата за первичку",
    "оплата за первичку $",
    "возврат",
    "возвращенные",
    "возвращенные средства",
    "возвращённые",
    "returned", "paid", "payers",
}

SALES_TOKENS = {
    "продажи",
    "продажи первичка",
    "продажи первичка $",
    "primary",
    "initial sales",
    "sales from primary",
}

COLD_TOKENS = {"холодная", "холодная база", "cold numbers", "cold"}

# Блоки, которые нужно полностью исключать из сводки по заголовку блока
# (не учитываются ни в одной метрике).
#   "dubai target" — на листе "Сводка CocoAge": лиды без бюджета, в общую сводку не идёт.
EXCLUDE_BLOCK_TITLE_TOKENS = {"dubai target"}

_WORD = r"a-z0-9а-я"


def _contains_any(v: Any, tokens: set) -> bool:
    s = _norm_h(v)
    for t in tokens:
        tn = _norm_base(str(t))
        if not tn:
            continue
        # короткие токены проверяем по "границам слова"
        if len(tn) <= 3 and re.fullmatch(rf"[{_WORD}]+", tn):
            if re.search(rf"(^|[^{_WORD}]){re.escape(tn)}([^{_WORD}]|$)", s):
                return True
        else:
            if tn in s:
                return True
    return False


# =========================
# Discovery model
# =========================
@dataclass
class TargetBlock:
    title: str
    start_col: str
    date: str
    leads: str = ""
    bookings_all: str = ""
    bookings_ai: str = ""
    bookings_op: str = ""
    arrived: str = ""
    budget: str = ""


@dataclass
class PairBlock:
    date_col: str
    value_col: str
    kind: str  # payers / sales / other_arrivals


@dataclass
class SheetDiscovery:
    sheet: str
    scan_start: str
    scan_end_requested: str
    scan_end_clamped: str
    header_row: int
    title_row: int
    work_row_start: int
    work_row_end: int
    target_blocks: List[TargetBlock]
    other_arrivals: List[PairBlock]
    payers: List[PairBlock]
    sales: List[PairBlock]
    excluded: List[Dict[str, str]]


def _find_header_row(grid: List[List[Any]], max_rows: int = 30) -> int:
    best_score = -1
    best_r = 3
    MIN_SCORE = 10
    for r0, row in enumerate(grid[:max_rows]):
        rown = [_norm_h(x) for x in row]
        if not any(v in DATE_TOKENS for v in rown):
            continue
        score = 0
        for v in rown:
            if v in DATE_TOKENS:
                score += 5
            if _contains_any(v, LEADS_TOKENS):
                score += 3
            if _contains_any(v, BUDGET_TOKENS):
                score += 3
            if _contains_any(v, ARRIVED_TOKENS):
                score += 2
            if _contains_any(v, BOOK_ALL_TOKENS):
                score += 2
            if _contains_any(v, BOOK_AI_TOKENS):
                score += 2
            if _contains_any(v, BOOK_OP_TOKENS):
                score += 2
            if _contains_any(v, PAYERS_TOKENS):
                score += 2
            if _contains_any(v, SALES_TOKENS):
                score += 2
        if score >= MIN_SCORE and score > best_score:
            best_score = score
            best_r = r0 + 1
    return best_r


def _detect_title_row(grid: List[List[Any]], header_row: int, date_positions: List[int]) -> int:
    candidates: List[int] = []
    if header_row - 1 >= 1:
        candidates.append(header_row - 1)
    candidates.append(1)
    best = 1
    best_score = -1
    for r in candidates:
        r0 = r - 1
        row = grid[r0] if 0 <= r0 < len(grid) else []
        score = 0
        for i in date_positions:
            if i < len(row):
                v = _norm_h(row[i])
                if v and v not in DATE_TOKENS:
                    score += 1
        if score > best_score:
            best_score = score
            best = r
    return best


def discover_sheet_from_grid(
    sheet_name: str,
    grid: List[List[Any]],
    scan_start_col: str,
    scan_end_requested: str,
    scan_end_clamped: str,
    work_rows: Tuple[int, int]
) -> SheetDiscovery:
    work_row_start, work_row_end = work_rows

    header_row = _find_header_row(grid, max_rows=min(30, len(grid)))
    hdr_r0 = header_row - 1
    hdr_raw = grid[hdr_r0] if 0 <= hdr_r0 < len(grid) else []
    hdr = [_norm_h(x) for x in hdr_raw]

    date_pos = [i for i, v in enumerate(hdr) if v in DATE_TOKENS]

    title_row = _detect_title_row(grid, header_row=header_row, date_positions=date_pos)
    title_r0 = title_row - 1
    titles_raw = grid[title_r0] if 0 <= title_r0 < len(grid) else []
    titles = [_s(x) for x in titles_raw]

    scan_start_idx = col_to_index(scan_start_col)

    def abs_col(offset0: int) -> str:
        return index_to_col(scan_start_idx + offset0)

    excluded: List[Dict[str, str]] = []

    # 1) Money pairs (верхние таблицы денег), чтобы их "Дата" не спутать с таргет-блоками
    other_arrivals: List[PairBlock] = []
    payers: List[PairBlock] = []
    sales: List[PairBlock] = []
    money_date_cols: set[int] = set()

    TOP_SCAN_ROWS_FOR_PAIRS = min(30, len(grid))
    for r0 in range(TOP_SCAN_ROWS_FOR_PAIRS):
        row = grid[r0]
        rown = [_norm_h(x) for x in row]
        for i in range(0, len(rown) - 1):
            if rown[i] not in DATE_TOKENS:
                continue
            h1 = rown[i + 1]
            if _contains_any(h1, ARRIVED_TOKENS):
                other_arrivals.append(PairBlock(abs_col(i), abs_col(i + 1), "other_arrivals"))
                money_date_cols.add(i)
                continue
            if _contains_any(h1, PAYERS_TOKENS):
                payers.append(PairBlock(abs_col(i), abs_col(i + 1), "payers"))
                money_date_cols.add(i)
                continue
            if _contains_any(h1, SALES_TOKENS):
                sales.append(PairBlock(abs_col(i), abs_col(i + 1), "sales"))
                money_date_cols.add(i)
                continue

    def _dedup(pairs: List[PairBlock]) -> List[PairBlock]:
        seen = set()
        out: List[PairBlock] = []
        for p in pairs:
            k = (p.date_col, p.value_col, p.kind)
            if k in seen:
                continue
            seen.add(k)
            out.append(p)
        return out

    other_arrivals = _dedup(other_arrivals)
    payers = _dedup(payers)
    sales = _dedup(sales)

    # 2) Target blocks по ЖЁСТКИМ оффсетам от "Дата"
    # date=0, leads=+1, bookings_all=+2, bookings_ai=+3, bookings_op=+4, arrived=+5, budget=+6
    target_blocks: List[TargetBlock] = []
    for i in date_pos:
        if i in money_date_cols:
            continue
        title = titles[i] if i < len(titles) else ""
        if _contains_any(title, COLD_TOKENS):
            excluded.append({"start_col": abs_col(i), "title": title, "reason": "cold_block_excluded"})
            continue
        if _contains_any(title, EXCLUDE_BLOCK_TITLE_TOKENS):
            excluded.append({"start_col": abs_col(i), "title": title, "reason": "excluded_by_title"})
            continue
        if i + 6 >= len(hdr):
            continue

        h_leads = hdr[i + 1]
        h_all = hdr[i + 2]
        h_ai = hdr[i + 3]
        h_op = hdr[i + 4]
        h_arr = hdr[i + 5]
        h_budget = hdr[i + 6]

        # строгая валидация структуры блока
        if not _contains_any(h_leads, LEADS_TOKENS):
            continue
        if not _contains_any(h_budget, BUDGET_TOKENS):
            continue
        if not (
            _contains_any(h_all, BOOK_ALL_TOKENS)
            or _contains_any(h_ai, BOOK_AI_TOKENS)
            or _contains_any(h_op, BOOK_OP_TOKENS)
            or _contains_any(h_arr, ARRIVED_TOKENS)
        ):
            continue

        start_col = abs_col(i)
        target_blocks.append(TargetBlock(
            title=title,
            start_col=start_col,
            date=start_col,
            leads=abs_col(i + 1),
            bookings_all=abs_col(i + 2),
            bookings_ai=abs_col(i + 3),
            bookings_op=abs_col(i + 4),
            arrived=abs_col(i + 5),
            budget=abs_col(i + 6),
        ))

    return SheetDiscovery(
        sheet=sheet_name,
        scan_start=scan_start_col,
        scan_end_requested=scan_end_requested,
        scan_end_clamped=scan_end_clamped,
        header_row=header_row,
        title_row=title_row,
        work_row_start=work_row_start,
        work_row_end=work_row_end,
        target_blocks=target_blocks,
        other_arrivals=other_arrivals,
        payers=payers,
        sales=sales,
        excluded=excluded,
    )


# =========================
# Overall formulas
# =========================
OVERALL_SHEET_DEFAULT = "Общая сводка все филиалы"
OVERALL_LABELS_DEFAULT = "B3:B60"  # с запасом; обход остановится на метке "KPI"

MODE_CELL = "$B$1"
FROM_CELL = "$C$1"
TO_CELL = "$D$1"
CURRENCY_TOGGLE = "$G$1"
DUBAI_RATE = "$E$1"   # USD -> AED
KZT_RATE = "$F$1"     # USD -> KZT
TASH_RATE = "$H$1"    # USD -> UZS


def _a1_col_range(sheet: str, col: str, r1: int, r2: int) -> str:
    return f"{quote_sheet(sheet)}!${col}${r1}:${col}${r2}"


def _safe_rate(cell: str) -> str:
    return f"ЕСЛИОШИБКА(ЕСЛИ(ИЛИ({cell}=\"\";{cell}=0);1;{cell});1)"


def _kzt_toggle_divider() -> str:
    """KZT: делим только когда переключатель = USD."""
    return f"ЕСЛИ({CURRENCY_TOGGLE}=\"USD\";{_safe_rate(KZT_RATE)};1)"


def _payers_divider(branch_kind: str) -> str:
    if branch_kind == "dubai":
        return _safe_rate(DUBAI_RATE)
    if branch_kind == "tashkent":
        return f"ЕСЛИ({CURRENCY_TOGGLE}=\"USD\";{_safe_rate(TASH_RATE)};1)"
    return _kzt_toggle_divider()


def _sales_divider(branch_kind: str) -> str:
    # CocoAge: продажи в AED -> USD (делим на курс AED/USD) всегда
    if branch_kind == "dubai":
        return _safe_rate(DUBAI_RATE)
    # Ташкент: продажи уже в USD -> НЕ делим
    if branch_kind == "tashkent":
        return "1"
    # KZ-филиалы и франшизы: продажи в KZT -> USD только когда переключатель = USD
    return _kzt_toggle_divider()


def _mode_norm_expr() -> str:
    return f'СТРОЧН(ПОДСТАВИТЬ(СЖПРОБЕЛЫ({MODE_CELL});"ё";"е"))'


def _has_expr(substr: str, norm_mode_expr: str) -> str:
    return f'ЕЧИСЛО(ЕСЛИОШИБКА(ПОИСК("{substr}";{norm_mode_expr});""))'


def _safe_date_expr(cell: str) -> str:
    return f"ЕСЛИ(ЕЧИСЛО({cell});{cell};ДАТАЗНАЧ({cell}))"


def _safe_to_date_expr() -> str:
    d_from = _safe_date_expr(FROM_CELL)
    d_to_raw = _safe_date_expr(TO_CELL)
    return f'ЕСЛИ(ИЛИ({TO_CELL}="";{TO_CELL}=0);{d_from};{d_to_raw})'


def make_mode_formula(ranges: List[Tuple[str, str]]) -> str:
    if not ranges:
        return "=0"
    vals = [v for (_, v) in ranges]
    sum_all = "СУММ(" + ";".join(vals) + ")"

    d_from = _safe_date_expr(FROM_CELL)
    d_to = _safe_to_date_expr()

    sum_one = "+".join([
        f'СУММЕСЛИМН({vr};{dr};">="&{d_from};{dr};"<"&({d_from}+1))'
        for (dr, vr) in ranges
    ])
    sum_between = "+".join([
        f'СУММЕСЛИМН({vr};{dr};">="&{d_from};{dr};"<"&({d_to}+1))'
        for (dr, vr) in ranges
    ])

    m = _mode_norm_expr()
    is_all = _has_expr("весь", m)
    is_single = _has_expr("определ", m)
    is_range = f"ИЛИ({_has_expr('диапаз', m)};{_has_expr('интервал', m)})"

    return (
        f'=ЕСЛИОШИБКА('
        f'ЕСЛИ({is_all};{sum_all};'
        f'ЕСЛИ({is_single};{sum_one};'
        f'ЕСЛИ({is_range};{sum_between};{sum_all})))'
        f';0)'
    )


def make_payers_formula(ranges: List[Tuple[str, str]], divider: str) -> str:
    if not ranges:
        return "=0"
    mode_expr = make_mode_formula(ranges)
    body = mode_expr[1:] if mode_expr.startswith("=") else mode_expr
    if divider == "1":
        return f"=ЕСЛИОШИБКА({body};0)"
    return f"=ЕСЛИОШИБКА(({body})/{divider};0)"


def make_sales_formula(ranges: List[Tuple[str, str]], divider: str) -> str:
    if not ranges:
        return "=0"
    mode_expr = make_mode_formula(ranges)
    body = mode_expr[1:] if mode_expr.startswith("=") else mode_expr
    if divider == "1":
        return f"=ЕСЛИОШИБКА({body};0)"
    return f"=ЕСЛИОШИБКА(({body})/{divider};0)"


def _build_ranges_for_metric(d: SheetDiscovery, metric: str) -> List[Tuple[str, str]]:
    sr, er = d.work_row_start, d.work_row_end
    out: List[Tuple[str, str]] = []

    def add(date_col: str, val_col: str):
        if not date_col or not val_col:
            return
        out.append((_a1_col_range(d.sheet, date_col, sr, er), _a1_col_range(d.sheet, val_col, sr, er)))

    if metric in ("budget", "leads", "arrived", "bookings_all", "bookings_ai", "bookings_op"):
        for tb in d.target_blocks:
            if metric == "budget":
                add(tb.date, tb.budget)
            elif metric == "leads":
                add(tb.date, tb.leads)
            elif metric == "arrived":
                add(tb.date, tb.arrived)
            elif metric == "bookings_all":
                add(tb.date, tb.bookings_all)
            elif metric == "bookings_ai":
                add(tb.date, tb.bookings_ai)
            elif metric == "bookings_op":
                add(tb.date, tb.bookings_op)
        if metric == "arrived":
            for p in d.other_arrivals:
                add(p.date_col, p.value_col)

    elif metric == "payers":
        for p in d.payers:
            add(p.date_col, p.value_col)

    elif metric == "sales":
        for p in d.sales:
            add(p.date_col, p.value_col)

    return out


def _derived_formulas(row: int) -> Dict[str, str]:
    return {
        "I": f"=ЕСЛИОШИБКА(C{row}/D{row};0)",
        "J": f"=ЕСЛИОШИБКА(C{row}/E{row};0)",
    }


# =========================
# EXACT mapping label -> sheet
# =========================
# Метки в overall!B (актуальный формат, июнь 2026):
#   ТАБЛИЦА "Филиал":
#     Филиал Кедма Астана
#     Филиал Кедма Алматы
#     Филиал Опатра Алматы
#     Филиал Опатра Айви Алматы
#     Филиал Опатра Ташкент
#     Опатра Ташкент Айви
#     Филиал Фэшн Астана
#     Филиал КокоЭйдж Дубаи
#     Айви Кедма Алматы
#     Айви Астана Фэшн
#     Филиал Нейротек
#     Vitally Life Астана
#     Vitally Life Алматы Рамзан
#   ТАБЛИЦА "Франшиза":
#     Франшиза VitalyLife Алматы
#     Франшиза VitalyLife Актау
#     Франшиза Vitally Life Усть-Каменогорск
LABEL_TO_SVODKA_RAW: Dict[str, str] = {
    # ---- Филиалы ----
    "Филиал Кедма Астана": "Сводка Кедма Астана",
    "Филиал Кедма Алматы": "Сводка Кедма Алматы",
    "Филиал Опатра Алматы": "Сводка Опатра Алматы",
    "Филиал Опатра Айви Алматы": "Сводка Опатра Айви Алматы",
    "Филиал Опатра Ташкент": "Сводка Опатра Ташкент",
    "Опатра Ташкент Айви": "Сводка Опатра Ташкент Айви",
    "Филиал Фэшн Астана": "Сводка Фэшн Астана",
    "Филиал КокоЭйдж Дубаи": "Сводка CocoAge",
    "Айви Кедма Алматы": "Сводка Айви Алматы",
    "Айви Астана Фэшн": "Сводка Айви Астана Фэшн",
    "Филиал Нейротек": "Сводка Нейротек",
    "Vitally Life Астана": "Сводка Vitally Life Астана",
    "Vitally Life Алматы Рамзан": "Сводка Vitally Life Алматы Рамзан",
    # ---- Франшизы ----
    "Франшиза VitalyLife Алматы": "Сводка Vitally Life Алматы Вадим",
    "Франшиза VitalyLife Актау": "Сводка Vitally Life Актау",
    "Франшиза Vitally Life Усть-Каменогорск": "Сводка Vitally Life Усть-Каменогорск",
}

LABEL_TO_SVODKA: Dict[str, str] = {_norm_label_key(k): v for k, v in LABEL_TO_SVODKA_RAW.items()}

# тип валюты для делителя платников/продаж
PAYERS_KIND_BY_SHEET = {
    "Сводка Опатра Ташкент": "tashkent",
    "Сводка Опатра Ташкент Айви": "tashkent",
    "Сводка CocoAge": "dubai",
    # остальные kz
}


def _infer_branch_kind_by_sheet(sheet_name: str) -> str:
    return PAYERS_KIND_BY_SHEET.get(sheet_name, "kz")


# =========================
# Patch generation
# =========================
@dataclass
class PatchResult:
    overall_sheet: str
    labels_range: str
    rows_to_update: List[int]
    sections: List[Tuple[int, int]]  # [(section_start_row, total_row), ...]
    updates_count: int
    discoveries: Dict[str, SheetDiscovery]
    row_map: Dict[int, str]


def generate_patch(
    spreadsheet_id: str,
    overall_sheet: str,
    overall_labels_range: str,
    work_rows: Tuple[int, int],
    scan: str,
    write_blockmap: Optional[str],
    only_blockmap: bool,
    verbose: bool = True,
) -> Tuple[Optional[List[Tuple[str, List[List[str]]]]], PatchResult]:
    service = _get_service()
    props = _get_sheets_properties(service, spreadsheet_id)

    scan_start, scan_end_req = parse_scan(scan)

    ranges: List[str] = []
    labels_a1 = f"{quote_sheet(overall_sheet)}!{overall_labels_range}"
    ranges.append(labels_a1)

    wanted_sheets = sorted(set(LABEL_TO_SVODKA.values()))
    HEADER_SCAN_ROWS = 30
    sheet_scan_range: Dict[str, str] = {}
    for sh in wanted_sheets:
        if sh not in props:
            if verbose:
                print(f"[WARN] sheet not found in spreadsheet: {sh}")
            continue
        col_count = int(props[sh].get("colCount") or 0)
        if col_count <= 0:
            continue
        max_col = index_to_col(col_count)
        end_idx = min(col_to_index(scan_end_req), col_to_index(max_col))
        end_col = index_to_col(end_idx)
        start_idx = col_to_index(scan_start)
        if start_idx > end_idx:
            continue
        a1 = f"{quote_sheet(sh)}!{scan_start}1:{end_col}{HEADER_SCAN_ROWS}"
        ranges.append(a1)
        sheet_scan_range[sh] = a1

    data = batch_get_formatted(service, spreadsheet_id, ranges)

    labels_vals = data.get(labels_a1, [])
    labels: List[str] = [(_clean_label(row[0]) if row else "") for row in labels_vals]

    start_row, _ = _parse_a1_row_range(overall_labels_range)

    # === Разбор меток на СЕКЦИИ ===
    # Каждая секция = непрерывный набор строк-филиалов, завершающийся строкой "Итог:".
    # Обход останавливается на метке "KPI" (всё ниже не трогаем).
    rows_to_update: List[int] = []
    row_map: Dict[int, str] = {}
    sections: List[Tuple[int, int]] = []
    section_start: Optional[int] = None

    for i, lab in enumerate(labels):
        row_num = start_row + i
        key = _norm_label_key(lab)

        # стоп на KPI-блоке
        if key in STOP_LABEL_KEYS:
            break

        if not lab:
            continue

        # "Итог:" -> закрываем текущую секцию
        if key.startswith("итог"):
            if section_start is not None:
                sections.append((section_start, row_num))
                section_start = None
            continue

        # служебные строки / заголовки секций
        if key in IGNORE_LABEL_KEYS:
            continue

        # инстаграм строки не трогаем (инста отдельным скриптом)
        if _is_instagram_label(lab):
            continue

        sh = LABEL_TO_SVODKA.get(key)
        if not sh:
            if verbose:
                print(f"[WARN] label not mapped (skip): '{lab}' (norm='{key}')")
            continue

        if section_start is None:
            section_start = row_num
        rows_to_update.append(row_num)
        row_map[row_num] = sh

    if verbose:
        print(f"[OVERALL] spreadsheet_id={spreadsheet_id}")
        print(f"[OVERALL] sheet={overall_sheet} labels={overall_labels_range} work_rows={work_rows[0]}:{work_rows[1]} scan={scan}")
        print(f"[OVERALL] rows_to_update={len(rows_to_update)} sections={sections}")

    # discoveries
    discoveries: Dict[str, SheetDiscovery] = {}
    for sh in sorted(set(row_map.values())):
        a1 = sheet_scan_range.get(sh)
        if not a1:
            continue
        grid = data.get(a1, [])
        m = re.match(r"^'(.+)'!(\w+)1:(\w+)\d+$", a1)
        scan_end_clamped = scan_end_req
        if m:
            scan_end_clamped = m.group(3)

        d = discover_sheet_from_grid(
            sheet_name=sh,
            grid=grid,
            scan_start_col=scan_start,
            scan_end_requested=scan_end_req,
            scan_end_clamped=scan_end_clamped,
            work_rows=work_rows,
        )
        discoveries[sh] = d
        if verbose:
            print(
                f"[DISCOVERY] {sh}: header_row={d.header_row} title_row={d.title_row} "
                f"targets={len(d.target_blocks)} payers={len(d.payers)} sales={len(d.sales)} other_arrivals={len(d.other_arrivals)}"
            )

    if write_blockmap:
        blockmap_obj = {
            "spreadsheet_id": spreadsheet_id,
            "overall_sheet": overall_sheet,
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "work_rows": {"start": work_rows[0], "end": work_rows[1]},
            "scan": {"start": scan_start, "end": scan_end_req},
            "sections": [{"start": s, "total": t} for (s, t) in sections],
            "sheets": {},
        }
        for sh, d in discoveries.items():
            blockmap_obj["sheets"][sh] = {
                "sheet": d.sheet,
                "scan_start": d.scan_start,
                "scan_end_requested": d.scan_end_requested,
                "scan_end_clamped": d.scan_end_clamped,
                "header_row": d.header_row,
                "title_row": d.title_row,
                "work_rows": {"start": d.work_row_start, "end": d.work_row_end},
                "target_blocks": [asdict(tb) for tb in d.target_blocks],
                "other_arrivals": [asdict(p) for p in d.other_arrivals],
                "payers": [asdict(p) for p in d.payers],
                "sales": [asdict(p) for p in d.sales],
                "excluded": d.excluded,
            }
        os.makedirs(os.path.dirname(write_blockmap) or ".", exist_ok=True)
        with open(write_blockmap, "w", encoding="utf-8") as f:
            json.dump(blockmap_obj, f, ensure_ascii=False, indent=2)
        if verbose:
            print(f"[BLOCKMAP] written: {write_blockmap}")

    if only_blockmap:
        res = PatchResult(
            overall_sheet=overall_sheet,
            labels_range=overall_labels_range,
            rows_to_update=rows_to_update,
            sections=sections,
            updates_count=0,
            discoveries=discoveries,
            row_map=row_map,
        )
        return None, res

    updates: List[Tuple[str, List[List[str]]]] = []

    # филиалы
    for row in rows_to_update:
        sh = row_map.get(row)
        if not sh:
            continue
        d = discoveries.get(sh)

        # если discovery не получился — зануляем, чтобы не оставались старые неверные формулы
        if not d:
            budget_f = "=0"
            leads_f = "=0"
            arrived_f = "=0"
            book_all_f = "=0"
            book_ai_f = "=0"
            book_op_f = "=0"
            payers_f = "=0"
            sales_f = "=0"
        else:
            kind = _infer_branch_kind_by_sheet(sh)
            budget_f = make_mode_formula(_build_ranges_for_metric(d, "budget"))
            leads_f = make_mode_formula(_build_ranges_for_metric(d, "leads"))
            arrived_f = make_mode_formula(_build_ranges_for_metric(d, "arrived"))
            book_all_f = make_mode_formula(_build_ranges_for_metric(d, "bookings_all"))
            book_ai_f = make_mode_formula(_build_ranges_for_metric(d, "bookings_ai"))
            book_op_f = make_mode_formula(_build_ranges_for_metric(d, "bookings_op"))
            payers_f = make_payers_formula(
                _build_ranges_for_metric(d, "payers"),
                _payers_divider(kind)
            )
            sales_f = make_sales_formula(
                _build_ranges_for_metric(d, "sales"),
                _sales_divider(kind)
            )

        der = _derived_formulas(row)
        row_values = [
            budget_f,    # C
            leads_f,     # D
            arrived_f,   # E
            book_all_f,  # F
            book_ai_f,   # G
            book_op_f,   # H
            der["I"],    # I
            der["J"],    # J
            payers_f,    # K
            sales_f,     # L
        ]
        rng = f"{quote_sheet(overall_sheet)}!C{row}:L{row}"
        updates.append((rng, [row_values]))

    # Итог по каждой секции: непрерывный диапазон section_start..(total_row-1)
    for (sec_start, total_row) in sections:
        if total_row <= sec_start:
            continue
        r1 = sec_start
        r2 = total_row - 1

        def sum_range(col: str, _r1=r1, _r2=r2) -> str:
            return f"=СУММ({col}{_r1}:{col}{_r2})"

        der_t = _derived_formulas(total_row)
        total_values = [
            sum_range("C"),
            sum_range("D"),
            sum_range("E"),
            sum_range("F"),
            sum_range("G"),
            sum_range("H"),
            der_t["I"],
            der_t["J"],
            sum_range("K"),
            sum_range("L"),
        ]
        total_rng = f"{quote_sheet(overall_sheet)}!C{total_row}:L{total_row}"
        updates.append((total_rng, [total_values]))

    res = PatchResult(
        overall_sheet=overall_sheet,
        labels_range=overall_labels_range,
        rows_to_update=rows_to_update,
        sections=sections,
        updates_count=len(updates),
        discoveries=discoveries,
        row_map=row_map,
    )
    return updates, res


def apply_updates(spreadsheet_id: str, updates: List[Tuple[str, List[List[str]]]], dry_run: bool, verbose: bool = True) -> None:
    if verbose:
        print(f"[PATCH] updates={len(updates)} ranges")
    if dry_run:
        for r, _ in updates:
            print(f"  - {r}")
        return
    service = _get_service()
    batch_update_user_entered(service, spreadsheet_id, updates)
    if verbose:
        print("[APPLY] done")


# =========================
# CLI
# =========================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet_id", required=True)
    ap.add_argument("--work_rows", default="4:33")   # июнь / 30 дней -> 4..33
    ap.add_argument("--scan", default="B:ET")
    ap.add_argument("--overall_sheet", default=OVERALL_SHEET_DEFAULT)
    ap.add_argument("--overall_labels_range", default=OVERALL_LABELS_DEFAULT)
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--dry_run", action="store_true")
    ap.add_argument("--quiet", action="store_true")
    ap.add_argument("--write_blockmap", default=None)
    ap.add_argument("--only_blockmap", action="store_true")
    args = ap.parse_args()

    verbose = (not args.quiet)
    work_rows = parse_work_rows(args.work_rows)

    updates, _res = generate_patch(
        spreadsheet_id=args.spreadsheet_id,
        overall_sheet=args.overall_sheet,
        overall_labels_range=args.overall_labels_range,
        work_rows=work_rows,
        scan=args.scan,
        write_blockmap=args.write_blockmap,
        only_blockmap=args.only_blockmap,
        verbose=verbose,
    )

    if args.only_blockmap:
        if verbose:
            print("[DONE] only_blockmap")
        return

    if not updates:
        if verbose:
            print("[PATCH] empty")
        return

    apply_updates(args.spreadsheet_id, updates, dry_run=args.dry_run or (not args.apply), verbose=verbose)


if __name__ == "__main__":
    main()
