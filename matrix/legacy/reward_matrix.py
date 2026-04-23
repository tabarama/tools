# -*- coding: utf-8 -*-
import argparse
import re
from pathlib import Path
from typing import Tuple, Optional, Any, Dict, List

import pandas as pd

# -------------------- DWH --------------------
DSN_NAME = "dwh_b2pos"
DWH_SCHEMA_TABLE = "mart_bi.excel_page_plan_kv_2"
DWH_DATE_COL = "dt_auth"   # поле даты в DWH для отбора месяца

# ---------- файлы/листы по умолчанию ----------
DEFAULT_DECK = r"C:\Users\r.taibov\Desktop\Engineering\FinAnalytics\1_deck_and_scenario\audit_deck.xlsm" # файл с наборами матриц (каждый лист — отдельная матрица)
DEFAULT_DATA = "reward_compression.xlsx" # файл, куда пишем Data + ScenarioResults
DATA_SHEET   = "Data"
AGRM_SHEET   = "tariff_agreements"
SCENARIO_SHEET = "ScenarioResults"
OUT_COLUMN   = "reward_matrix"

# -------------------- единый справочник полей --------------------
# matrix — одно каноничное имя, которое используется в tariff_matrix.parameter
FIELD_MAP: Dict[str, Dict[str, str]] = {
    "idcredit":        {"dwh": "idcredit",        "data": "idcredit",        "matrix": "idcredit"},
    "date":            {"dwh": "dt_auth",         "data": "date",            "matrix": "date"},
    "proc_type":       {"dwh": "proc_type",       "data": "proc_type",       "matrix": "proc_type"},
    "bank_name":       {"dwh": "bank",            "data": "bank_name",       "matrix": "bank_name"},
    "idstock":         {"dwh": "idstock",         "data": "idstock",         "matrix": "idstock"},
    "stock_code":     {"dwh": "stock_code",      "data": "stock_code",     "matrix": "stock_code"},
    "stock":           {"dwh": "stock",           "data": "stock",           "matrix": "stock"},
    "rate":            {"dwh": "rate_perc",       "data": "rate",            "matrix": "rate"},
    "term":            {"dwh": "term",            "data": "term",            "matrix": "term"},
    "sumcredit":       {"dwh": "sumcredit",       "data": "sumcredit",       "matrix": "sumcredit"},
    "agreement_name":  {"dwh": "agreement_name",  "data": "agreement_name",  "matrix": "agreement_name"},
    "real_reward":     {"dwh": "reward",      "data": "real_reward",     "matrix": "real_reward"},
}

COL_ID       = FIELD_MAP["idcredit"]["data"]
COL_BANK     = FIELD_MAP["bank_name"]["data"]
COL_PROC     = FIELD_MAP["proc_type"]["data"]
COL_AGR      = FIELD_MAP["agreement_name"]["data"]
COL_FALLBACK = FIELD_MAP["real_reward"]["data"]

# ---------- обязательные столбцы в матрице ----------
TM_BANK   = "bank_name"
TM_PROC   = "proc_type"
TM_CODE   = "tariff_code"
TM_GROUP  = "group"
TM_PARAM  = "parameter"
TM_OP     = "condition_type"
TM_VAL    = "value"
TM_REWARD = "reward_value"

# ---------- обязательные столбцы в tariff_agreements ----------
MAP_BANK  = "bank_name"
MAP_PROC  = "proc_type"
MAP_AGR   = "agreement_name"
MAP_CODE  = "tariff_code"

NBSP = "\u00A0"


def build_sql_and_maps(field_map: Dict[str, Dict[str, str]]) -> Tuple[str, Dict[str, str], Dict[str, str]]:
    dwh_cols: List[str] = []
    rename_map: Dict[str, str] = {}
    matrix_to_data: Dict[str, str] = {}
    for _, spec in field_map.items():
        dwh_cols.append(spec["dwh"])
        rename_map[spec["dwh"]] = spec["data"]
        matrix_to_data[spec["matrix"]] = spec["data"]
    return ", ".join(dwh_cols), rename_map, matrix_to_data


SQL_SELECT_LIST, DWH_RENAME_MAP, MATRIX_TO_DATA = build_sql_and_maps(FIELD_MAP)


# ---------- нормализация ----------
def clean_text(s: Any) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\t", " ").replace("\r", " ").replace("\n", " ").replace(NBSP, " ")
    s = " ".join(s.split())
    s = s.replace(" .", ".").replace(". ", ".")
    s = s.strip().lower().replace("ё", "е")
    return s


def norm_proc(x: Any) -> str:
    s = clean_text(x)
    if s in ("0", "inside"):
        return "inside"
    if s in ("1", "outside"):
        return "outside"
    return s


def norm_agr(x: Any) -> str:
    s = clean_text(x)
    s = re.sub(r"^[\s0-9._\-№]+", "", s)
    return s.strip()


def make_rules_key(bank_norm: str, proc_norm: str, code_norm: str) -> str:
    return f"{bank_norm}|{proc_norm}|{code_norm}"


# ---------- сравнения ----------
def to_number_safe(v: Any) -> Optional[float]:
    if v is None:
        return None
    s = str(v).strip().replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def try_parse_iso_date(v: Any) -> Optional[pd.Timestamp]:
    s = str(v).strip()
    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        try:
            return pd.to_datetime(s, format="%Y-%m-%d")
        except Exception:
            return None
    try:
        return pd.to_datetime(v)
    except Exception:
        return None


def sql_like_match(cell_val: Any, pattern: str) -> bool:
    text = "" if cell_val is None else str(cell_val).strip().lower().replace("ё", "е")
    pat = str(pattern).strip().lower().replace("ё", "е")
    pat_esc = re.escape(pat).replace(r"\%", ".*").replace(r"\_", ".")
    return re.search("^" + pat_esc + "$", text) is not None


def compare_values(cell_val: Any, op: str, val_text: str) -> bool:
    op = (op or "").strip()

    v_num = to_number_safe(val_text)
    c_num = to_number_safe(cell_val)
    if v_num is not None and c_num is not None:
        if op == "==": return c_num == v_num
        if op in ("!=", "<>"): return c_num != v_num
        if op == "<": return c_num < v_num
        if op == "<=": return c_num <= v_num
        if op == ">": return c_num > v_num
        if op == ">=": return c_num >= v_num
        return False

    v_dt = try_parse_iso_date(val_text)
    c_dt = try_parse_iso_date(cell_val)
    if v_dt is not None and c_dt is not None:
        if op == "==": return c_dt == v_dt
        if op in ("!=", "<>"): return c_dt != v_dt
        if op == "<": return c_dt < v_dt
        if op == "<=": return c_dt <= v_dt
        if op == ">": return c_dt > v_dt
        if op == ">=": return c_dt >= v_dt
        return False

    s_cell = "" if cell_val is None else str(cell_val).strip()
    s_val = str(val_text).strip()
    if op == "==": return s_cell == s_val
    if op in ("!=", "<>"): return s_cell != s_val
    return False


def parse_bounds(bounds: str) -> Tuple[Optional[str], Optional[str], bool, bool]:
    """
    "10;20" -> default [10;20)
    "[10;20)", "(10;20]", "[10;20]", "(10;20)"
    "10;" or ";20"
    """
    s = str(bounds).strip()
    left_incl = True
    right_incl = False
    if s and s[0] in "([":
        left_incl = (s[0] == "[")
        s = s[1:].strip()
    if s and s[-1] in ")]":
        right_incl = (s[-1] == "]")
        s = s[:-1].strip()
    parts = [p.strip() for p in s.split(";")]
    lo = parts[0] if len(parts) > 0 and parts[0] != "" else None
    hi = parts[1] if len(parts) > 1 and parts[1] != "" else None
    return lo, hi, left_incl, right_incl


def between_ok(cell_val: Any, bounds: str) -> bool:
    lo, hi, left_incl, right_incl = parse_bounds(bounds)
    if lo is not None and not compare_values(cell_val, ">=" if left_incl else ">", lo):
        return False
    if hi is not None and not compare_values(cell_val, "<=" if right_incl else "<", hi):
        return False
    return True


def eval_cond(cell_val: Any, op: str, val_text: str) -> bool:
    """
    Supported:
      =, ==, !=, <>, <, <=, >, >=
      between
      in / not_in      (list via ;)
      like / not_like  (patterns via ;)
    """
    op = (op or "").strip().lower()

    if op == "=":
        op = "=="

    if op == "between":
        return between_ok(cell_val, val_text)

    if op in ("in", "not_in"):
        vals = [v.strip() for v in str(val_text).split(";") if v.strip()]
        if not vals:
            return False
        match = any(compare_values(cell_val, "==", v) for v in vals)
        return match if op == "in" else not match

    if op in ("like", "not_like"):
        pats = [p.strip() for p in str(val_text).split(";") if p.strip()]
        if not pats:
            return False
        match = any(sql_like_match(cell_val, p) for p in pats)
        return match if op == "like" else not match

    if op in ("==", "!=", "<>", "<", "<=", ">", ">="):
        return compare_values(cell_val, op, val_text)

    return False


# ---------- индексы правил/маппинга ----------
def build_rules_index(rules: pd.DataFrame) -> Dict[str, Dict]:
    t = rules.copy()
    t["_bank"] = t[TM_BANK].apply(clean_text)
    t["_proc"] = t[TM_PROC].apply(norm_proc)
    t["_code"] = t[TM_CODE].apply(clean_text)
    t["_key"]  = t.apply(lambda r: make_rules_key(r["_bank"], r["_proc"], r["_code"]), axis=1)

    for col in [TM_GROUP, TM_PARAM, TM_OP, TM_VAL, TM_REWARD]:
        if col in t.columns:
            t[col] = t[col].astype(str)

    index: Dict[str, Dict] = {}
    for key, dfk in t.groupby("_key", sort=False):
        groups: Dict[str, Dict] = {}
        for grp, dfg in dfk.groupby(TM_GROUP, sort=False):
            rows = []
            min_reward = None
            for _, rr in dfg.iterrows():
                param = str(rr[TM_PARAM]).strip()
                op    = str(rr[TM_OP]).strip()
                val   = str(rr[TM_VAL]).strip()
                rwd   = to_number_safe(rr[TM_REWARD])
                rows.append((param, op, val, rwd))
                if rwd is not None:
                    min_reward = rwd if min_reward is None else min(min_reward, rwd)
            groups[str(grp)] = {"rows": rows, "min_reward": min_reward}
        index[key] = groups
    return index


def build_agreement_index(agr: pd.DataFrame) -> Dict[str, Dict]:
    t = agr.copy()
    t["_bank"] = t[MAP_BANK].apply(clean_text)
    t["_proc"] = t[MAP_PROC].apply(norm_proc)
    t["_agr"]  = t[MAP_AGR].apply(norm_agr)
    t["_code"] = t[MAP_CODE].apply(clean_text)

    idx: Dict[str, Dict] = {}
    for _, r in t.iterrows():
        b, p, a, c = r["_bank"], r["_proc"], r["_agr"], r["_code"]
        idx.setdefault(b, {}).setdefault(p, {})[a] = c
    return idx


# ---------- выбор листов-сценариев ----------
def get_scenario_sheets(deck_path: Path, explicit: Optional[str]) -> List[str]:
    wb = pd.ExcelFile(deck_path, engine="openpyxl")
    all_sheets = wb.sheet_names
    required_cols = {TM_BANK, TM_PROC, TM_CODE, TM_GROUP, TM_PARAM, TM_OP, TM_VAL, TM_REWARD}

    valid_sheets: List[str] = []
    for s in all_sheets:
        if s == AGRM_SHEET:
            continue
        try:
            df0 = wb.parse(s, nrows=0)
        except Exception:
            continue
        if required_cols.issubset(set(df0.columns)):
            valid_sheets.append(s)

    if not valid_sheets:
        raise ValueError(
            f"В '{deck_path.name}' не найдено ни одного листа с полным набором колонок матрицы "
            f"(ожидались: {sorted(required_cols)})."
        )

    if explicit:
        wanted = [s.strip() for s in explicit.split(",") if s.strip()]
        missing = [s for s in wanted if s not in valid_sheets]
        if missing:
            raise ValueError(f"Листы {missing} не подходят как матрицы/отсутствуют. Доступные: {valid_sheets}")
        return wanted

    print(f"Найдено листов-матриц в '{deck_path.name}':")
    for i, s in enumerate(valid_sheets, 1):
        print(f"{i}. {s}")

    while True:
        raw = input("Введи номера листов-матриц через запятую или '*' чтобы взять все: ").strip()
        if raw == "*":
            return valid_sheets

        nums = [p.strip() for p in raw.split(",") if p.strip()]
        if not nums:
            print("Нужно ввести хотя бы один номер или '*'.")
            continue

        try:
            idxs = [int(x) for x in nums]
        except ValueError:
            print("Допустимы только числа или '*'.")
            continue

        bad = [n for n in idxs if n < 1 or n > len(valid_sheets)]
        if bad:
            print(f"Некорректные номера: {bad}. Диапазон: 1..{len(valid_sheets)}.")
            continue

        return [valid_sheets[n - 1] for n in idxs]


# ---------- DWH (выбор месяца) ----------
def first_day_of_month(ts: pd.Timestamp) -> pd.Timestamp:
    ts = pd.Timestamp(ts)
    return ts.replace(day=1, hour=0, minute=0, second=0, microsecond=0)


def first_day_next_month(ts: pd.Timestamp) -> pd.Timestamp:
    ts = first_day_of_month(ts)
    if ts.month == 12:
        return ts.replace(year=ts.year + 1, month=1)
    return ts.replace(month=ts.month + 1)


def month_range_from_arg(month_arg: str, max_dt: pd.Timestamp) -> Tuple[pd.Timestamp, pd.Timestamp]:
    month_arg = (month_arg or "").strip().lower()

    if month_arg in ("", "latest"):
        start = first_day_of_month(max_dt)
        end = first_day_next_month(max_dt)
        return start, end

    m = re.fullmatch(r"(\d{4})-(\d{2})", month_arg)
    if not m:
        raise ValueError("Аргумент --month должен быть 'latest' или формата YYYY-MM (например 2026-02)")
    y = int(m.group(1))
    mo = int(m.group(2))
    if not (1 <= mo <= 12):
        raise ValueError("Месяц в --month должен быть 01..12")

    start = pd.Timestamp(year=y, month=mo, day=1)
    end = first_day_next_month(start)
    return start, end


def load_month_from_dwh(dsn: str, schema_table: str, select_list: str, date_col: str, month_arg: str) -> Tuple[pd.DataFrame, pd.Timestamp, pd.Timestamp]:
    try:
        import pyodbc  # noqa
    except Exception as e:
        raise RuntimeError("Не найден пакет 'pyodbc'. Установи: pip install pyodbc") from e

    import pyodbc

    conn_str = f"DSN={dsn};"
    with pyodbc.connect(conn_str) as conn:
        sql_max = f"SELECT MAX({date_col}) AS max_dt FROM {schema_table}"
        max_df = pd.read_sql(sql_max, conn)
        if max_df.empty or max_df.loc[0, "max_dt"] is None:
            raise RuntimeError(f"Не удалось получить MAX({date_col}) из {schema_table}")
        max_dt = pd.to_datetime(max_df.loc[0, "max_dt"])

        start, end = month_range_from_arg(month_arg, max_dt)

        sql_data = f"SELECT {select_list} FROM {schema_table} WHERE {date_col} >= ? AND {date_col} < ?"
        df = pd.read_sql(sql_data, conn, params=[start.to_pydatetime(), end.to_pydatetime()])

    return df, start, end


# ---------- проверка групп ----------
def evaluate_groups(data_row: pd.Series, groups: Dict, matrix_to_data: Dict[str, str]) -> Tuple[str, str, Optional[float]]:
    passed_rewards: List[float] = []
    passed_groups: List[str] = []

    best_satisfied = -1
    best_fail_param = None
    best_fail_group = None

    for grp_name, ginfo in groups.items():
        ok_all = True
        satisfied = 0
        fail_param_local = None

        for param, op, val, _ in ginfo["rows"]:
            col = matrix_to_data.get(param)
            if not col or col not in data_row.index:
                ok_all = False
                fail_param_local = param
                break

            if eval_cond(data_row[col], op, val):
                satisfied += 1
            else:
                ok_all = False
                fail_param_local = param
                break

        if ok_all:
            rwd = ginfo["min_reward"]
            passed_rewards.append(0.0 if rwd is None else float(rwd))
            passed_groups.append(grp_name)
        else:
            if satisfied > best_satisfied:
                best_satisfied = satisfied
                best_fail_param = fail_param_local
                best_fail_group = grp_name

    if passed_rewards:
        m = min(passed_rewards)
        g = passed_groups[passed_rewards.index(m)]
        return "OK", g, m

    param_name = str(best_fail_param) if best_fail_param is not None else "UNKNOWN"
    param_name = re.sub(r"\s+", "_", param_name)
    return f"NO_{param_name}", (str(best_fail_group) if best_fail_group else ""), None


# ---------- main ----------
def main():
    parser = argparse.ArgumentParser(
        description="Применение нескольких матриц (сценариев) из deck.xlsm к данным из DWH и запись в reward_compression.xlsx."
    )
    parser.add_argument("--deck", default=DEFAULT_DECK, help="файл с матрицами (deck.xlsm)")
    parser.add_argument("--data", default=DEFAULT_DATA, help="файл результата (reward_compression.xlsx)")
    parser.add_argument("--scenarios", default=None, help="имена листов-сценариев через запятую или интерактивный выбор")
    parser.add_argument("--dsn", default=DSN_NAME, help="ODBC DSN name")
    parser.add_argument("--table", default=DWH_SCHEMA_TABLE, help="schema.table in DWH")
    parser.add_argument("--date_col", default=DWH_DATE_COL, help="дата-колонка в DWH для отбора месяца")
    parser.add_argument("--month", default="latest", help="какой месяц брать: latest или YYYY-MM (например 2026-02)")
    args = parser.parse_args()

    deck_path = Path(args.deck)
    if not deck_path.is_absolute():
        deck_path = (Path(__file__).resolve().parent / args.deck).resolve()

    data_path = Path(__file__).with_name(args.data)

    if not deck_path.exists():
        raise FileNotFoundError(f"Не найден файл с матрицами: {deck_path}")

    # 1) выбираем сценарии
    scenario_sheets = get_scenario_sheets(deck_path, args.scenarios)
    print(f"Сценарии (листы) для расчёта: {scenario_sheets}")

    # 2) грузим данные из DWH за выбранный месяц
    print(f"Читаю DWH через DSN='{args.dsn}', table='{args.table}', month='{args.month}'...")
    df_dwh, m_start, m_end = load_month_from_dwh(args.dsn, args.table, SQL_SELECT_LIST, args.date_col, args.month)
    print(f"Загружен диапазон: [{m_start.date()} ; {m_end.date()})  строк: {len(df_dwh):,}".replace(",", " "))

    # 3) переименовываем в "data"-имена
    df = df_dwh.rename(columns=DWH_RENAME_MAP)

    # 4) real_reward делим на 100 (как договорились)
    df[COL_FALLBACK] = pd.to_numeric(
        df[COL_FALLBACK].astype(str).str.replace(" ", "").str.replace(",", "."),
        errors="coerce"
    ) / 100

    # проверки обязательных полей Data
    need = [
        FIELD_MAP["idcredit"]["data"],
        FIELD_MAP["bank_name"]["data"],
        FIELD_MAP["proc_type"]["data"],
        FIELD_MAP["agreement_name"]["data"],
        FIELD_MAP["real_reward"]["data"],
        FIELD_MAP["rate"]["data"],
        FIELD_MAP["term"]["data"],
        FIELD_MAP["sumcredit"]["data"],
        FIELD_MAP["date"]["data"],
    ]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise ValueError(f"После переименования нет обязательных колонок в Data: {missing}")

    # 5) читаем маппинг договоров
    wb_deck = pd.ExcelFile(deck_path, engine="openpyxl")
    agrm = wb_deck.parse(AGRM_SHEET)
    for c in [MAP_BANK, MAP_PROC, MAP_AGR, MAP_CODE]:
        if c not in agrm.columns:
            raise ValueError(f"В '{deck_path.name}' на листе '{AGRM_SHEET}' нет колонки '{c}'")

    agr_index = build_agreement_index(agrm)

    all_results: List[pd.DataFrame] = []

    # 6) цикл по сценариям
    for sheet_name in scenario_sheets:
        print(f"Считаем сценарий '{sheet_name}'...")

        rules = wb_deck.parse(sheet_name)
        for c in [TM_BANK, TM_PROC, TM_CODE, TM_GROUP, TM_PARAM, TM_OP, TM_VAL, TM_REWARD]:
            if c not in rules.columns:
                raise ValueError(f"В '{deck_path.name}' на листе '{sheet_name}' нет колонки '{c}'")

        rules_index = build_rules_index(rules)

        scen_status: List[str] = []
        scen_group: List[str]  = []
        scen_reward: List[Any] = []

        for _, row in df.iterrows():
            bank_norm = clean_text(row[COL_BANK])
            proc_norm = norm_proc(row[COL_PROC])
            agr_norm  = norm_agr(row[COL_AGR])

            st: str
            grp: str = ""
            rwd: Optional[float] = None

            if bank_norm not in agr_index:
                st = "NO_BANK"
            elif proc_norm not in agr_index[bank_norm]:
                st = "NO_PROC_TYPE"
            elif agr_norm not in agr_index[bank_norm][proc_norm]:
                st = "NO_AGREEMENT"
            else:
                code_norm = agr_index[bank_norm][proc_norm][agr_norm]
                key = make_rules_key(bank_norm, proc_norm, code_norm)
                groups = rules_index.get(key)
                if not groups:
                    st = "NO_AGREEMENT"
                else:
                    st_loc, grp_loc, rwd_loc = evaluate_groups(row, groups, MATRIX_TO_DATA)
                    st, grp, rwd = st_loc, grp_loc, rwd_loc

            if st in ("NO_BANK", "NO_PROC_TYPE", "NO_AGREEMENT"):
                reward_val = row[COL_FALLBACK]
            elif st == "OK":
                reward_val = rwd
            else:
                reward_val = "#N/A"

            scen_status.append(st)
            scen_group.append(grp)
            scen_reward.append(reward_val)

        df_scen = pd.DataFrame({
            "scenario_name": sheet_name,
            COL_ID: df[COL_ID],
            COL_BANK: df[COL_BANK],
            COL_PROC: df[COL_PROC],
            COL_AGR: df[COL_AGR],
            OUT_COLUMN: scen_reward,
            "matrix_status": scen_status,
            "matrix_group": scen_group,
        })

        rm_num = pd.to_numeric(pd.Series(scen_reward).astype(str).str.replace(" ", "").str.replace(",", "."),
                               errors="coerce")
        sum_num = pd.to_numeric(df["sumcredit"].astype(str).str.replace(" ", "").str.replace(",", "."),
                                errors="coerce")
        df_scen["reward_matrix_amount"] = rm_num * sum_num

        all_results.append(df_scen)

    res_all = pd.concat(all_results, ignore_index=True)

    # 7) пишем результат
    with pd.ExcelWriter(data_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=DATA_SHEET, index=False)
        res_all.to_excel(writer, sheet_name=SCENARIO_SHEET, index=False)

    print(f"Готово. Записано в '{data_path.name}': лист '{DATA_SHEET}' и лист '{SCENARIO_SHEET}'.")


if __name__ == "__main__":
    main()
