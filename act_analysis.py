# -*- coding: utf-8 -*-
"""
Анализ актов v2.9.9 (основная ветка: Общ_Сводная).
База: act_analysis_branch_DP v1.0 + учёт актов из Общей сводной без вхождения в Сводную ГП.
Добавлено: столбец G «DP», сдвиг выходных колонок C:V.
Устойчивость: обрабатывает пропуски в столбце B (пустые ячейки между актами).
Важно: акт отображается, если он есть в Общей сводной, даже если отсутствует в Сводной ГП.
Запуск:
    python act_analysis.py          — обновить цифры для текущего списка в B
    python act_analysis.py collect  — пересобрать список актов из сводных
"""

import sys
import time
import re
from pathlib import Path

# ------------------ настройки источников ------------------
GP_PATH = (
    r"\\vls.lan\ULSPG-DFS\ПТО\04. Завод СПГ_Рабочая общая"
    r"\02. Отчёты по исполнительной документации\Группа КМ\3. Генподряд"
    r"\!СВОДНЫЕ ГП КМ\Сводная_ГП_КМ_ПК"
)
OB_PATH = (
    r"\\vls.lan\ULSPG-DFS\ПТО\04. Завод СПГ_Рабочая общая"
    r"\02. Отчёты по исполнительной документации"
    r"\ОБЩАЯ Сводная таблица по статусу объекта СПГ КМ (по линиям)"
    r"\ОБЩАЯ Сводная таблица по статусу объекта СПГ КМ (ПК"
)

# ------------------ столбцы источников (0-based) ------------------
GP_J, GP_S, GP_W, GP_AB, GP_AI, GP_AM = 9, 18, 22, 27, 34, 38
GP_N = 13                                   # DP, столбец N источника
OB_J, OB_W, OB_AO = 9, 22, 40
MET_D, MET_G, MET_H, MET_J = 3, 6, 7, 9

SKIP_ROWS = 9
MET_SKIP_ROWS = 1

XL_UP = -4162
XL_MANUAL = -4135
LIGHT_BLUE = 173 + 216 * 256 + 230 * 65536

# ------------------ выходной блок ------------------
TOTAL_COLS = 20                              # C..V
LAST_OUTPUT = 22                             # V
LAST_TABLE = 28                              # AB — все столбцы с данными в строке 1

OGZ_PREFIXES = ("100б", "100н", "100холд")
SPECIAL_PREFIXES = OGZ_PREFIXES + ("100дс",)

VERSION = "2.9.9"

# ------------------ вспомогательные функции ------------------
def cell(row, i):
    return row[i] if i < len(row) else None

def txt(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else str(v)
    if isinstance(v, int):
        return str(v)
    s = str(v)
    s = (s.replace("\xa0", " ").replace("\t", " ")
         .replace("\r", " ").replace("\n", " "))
    while "  " in s:
        s = s.replace("  ", " ")
    return s.strip()

def num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = txt(v).replace(" ", "").replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None

def title_of(act):
    p1 = act.find("(")
    p2 = act.find(")")
    if p1 < 0 or p2 <= p1:
        return "-"
    return act[p1 + 1:p2]

def pc_of(ttl):
    head = ttl.split(".")[0]
    if ttl[:3] == "L03" or head in ("L3-02-010", "L3-12-030"):
        return "ПК1"
    if ttl[:2] == "L1" or ttl[:3] == "L01":
        return "ПК2"
    if ttl[:2] == "L3":
        return "ПК3"
    return ""

def sec_of(act):
    p = act.find("-КМ")
    if p < 0:
        return "-"
    left = act[:p]
    dot = left.find(".")
    if dot < 0:
        return "-"
    dash = left.find("-", dot + 1)
    if dash < 0:
        return "-"
    return left[dot + 1:dash]

def km_of(act):
    p = act.find("-27")
    if p < 0:
        return "-"
    return "КМ" + act[p + 4:p + 5]

def _exec_hit(comp, exec_str):
    if comp == exec_str:
        return True
    return (comp.startswith(exec_str) and len(comp) > len(exec_str)
            and not comp[len(exec_str)].isdigit())

def classify(st, exec_str):
    parts = [p.strip() for p in (st or "").split(";") if p.strip()]
    if not parts:
        return "out"
    low = [p.lower() for p in parts]
    cur = int(exec_str)
    if any(p.isdigit() and int(p) < cur for p in parts):
        return "past"
    if (any(p.startswith("100б") or p.startswith("100н") or p.startswith("100дс")
            or p.startswith("100холд") for p in low)):
        return "ogz"
    if any(p == "сварка" for p in low):
        return "svarka"
    if any(_exec_hit(p, exec_str) for p in parts):
        return "prot"
    return "out"

def had_past_exec(st_iter, exec_str):
    cur = int(exec_str)
    for st in st_iter:
        for p in st.split(";"):
            p = p.strip()
            if p.isdigit() and int(p) < cur:
                return True
    return False

def act_in_exec(st_iter, exec_str):
    for st in st_iter:
        for p in st.split(";"):
            if _exec_hit(p.strip(), exec_str):
                return True
    return False

def sort_key(act):
    ttl = title_of(act)
    sec = sec_of(act)
    km = km_of(act)
    sec_n = int(sec) if sec.isdigit() else 10 ** 9
    km_n = int(km[3:]) if km[:3] == "КМ" and km[3:].isdigit() else 10 ** 9
    return (pc_of(ttl), ttl, sec_n, km_n, act)

# ------------------ чтение источников ------------------
def read_gp(path, exec_str):
    """-> (agg: act->[sumS, ogz, svarka, prot, weld, ogz_area, cnt],
           sts: act->set, dps: act->set)"""
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(str(path))
    rows = wb.get_sheet_by_name("ГПВ КМ").to_python()
    agg, sts, dps = {}, {}, {}
    for row in rows[SKIP_ROWS:]:
        act = txt(cell(row, GP_AI))
        if not act:
            continue
        if not txt(cell(row, GP_J)):
            continue
        dp = txt(cell(row, GP_N))
        if dp:
            dps.setdefault(act, set()).add(dp)
        st = txt(cell(row, GP_AM))
        s = num(cell(row, GP_S))
        w = num(cell(row, GP_W))
        area = num(cell(row, GP_AB))
        a = agg.get(act)
        if a is None:
            a = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0]
            agg[act] = a
        if s is not None:
            a[0] += s
        cls = classify(st, exec_str)
        if cls == "ogz":
            a[1] += s if s else 0.0
            a[5] += area if area else 0.0
        elif cls == "svarka":
            a[2] += s if s else 0.0
        elif cls == "prot":
            a[3] += s if s else 0.0
        if w is not None:
            a[4] += w
        a[6] += 1
        if st:
            sts.setdefault(act, set()).add(st)
    return agg, sts, dps

def read_ob(path):
    """-> agg: act->[cnt, sumW]"""
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(str(path))
    rows = wb.get_sheet_by_name("Сводная КМ").to_python()
    agg = {}
    for row in rows[SKIP_ROWS:]:
        act = txt(cell(row, OB_AO))
        if not act:
            continue
        if not txt(cell(row, OB_J)):
            continue
        w = num(cell(row, OB_W))
        a = agg.get(act)
        if a is None:
            a = [0.0, 0.0]
            agg[act] = a
        a[0] += 1
        if w is not None:
            a[1] += w
    return agg

def read_metizy(path):
    """-> agg: act->[sum_G, sum_H, sum_J] (к освоению, ВПБ, по ИД)"""
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(str(path))
    rows = wb.get_sheet_by_name("Метизы").to_python()
    agg = {}
    for row in rows[MET_SKIP_ROWS:]:
        act = txt(cell(row, MET_D))
        if not act:
            continue
        g = num(cell(row, MET_G))
        h = num(cell(row, MET_H))
        j = num(cell(row, MET_J))
        a = agg.get(act)
        if a is None:
            a = [0.0, 0.0, 0.0]
            agg[act] = a
        if g is not None:
            a[0] += g
        if h is not None:
            a[1] += h
        if j is not None:
            a[2] += j
    return agg

# ------------------ основной скрипт ------------------
def main():
    import win32api
    import win32con
    import win32com.client as wc

    t0 = time.time()

    def die(msg, code=1):
        win32api.MessageBox(0, msg, "act_analysis_branch_DP", win32con.MB_ICONERROR)
        print(msg)
        return code

    collect = len(sys.argv) > 1 and sys.argv[1].strip().lower() == "collect"

    try:
        xl = wc.GetActiveObject("Excel.Application")
    except Exception:
        return die("Не найден запущенный Excel.\n"
                   "Открой книгу с анализом и запусти скрипт заново.")

    wbk = xl.ActiveWorkbook
    if wbk is None:
        return die("В Excel нет ни одной открытой книги.")
    ws = xl.ActiveSheet

    try:
        a1 = ws.Range("A1").Value
    except Exception:
        return die("Активный лист книги «%s» не похож на таблицу." % wbk.Name)

    exec_str = txt(a1)
    if not exec_str or not exec_str.replace(".", "", 1).isdigit():
        return die("В A1 листа «%s» книги «%s» нет номера выполнения."
                   % (ws.Name, wbk.Name))

    old_last = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row

    gp_by_pc, ob_by_pc, met_by_pc = {}, {}, {}
    warns = []

    def load_pc(pc):
        n = pc[2]
        t1 = time.time()
        try:
            gp_by_pc[pc] = read_gp(GP_PATH + n + ".xlsb", exec_str)
        except Exception as e:
            gp_by_pc[pc] = None
            warns.append("Нет источника ГП по %s (%s)." % (pc, e))
        try:
            ob_by_pc[pc] = read_ob(OB_PATH + n + ").xlsb")
        except Exception as e:
            ob_by_pc[pc] = None
            warns.append("Нет источника Общей по %s (%s)." % (pc, e))
        try:
            met_by_pc[pc] = read_metizy(GP_PATH + n + ".xlsb")
        except Exception as e:
            met_by_pc[pc] = None
            warns.append("Нет листа Метизы по %s (%s)." % (pc, e))
        print(pc, "read %.1fs" % (time.time() - t1))

    if collect:
        for pc in ("ПК1", "ПК2", "ПК3"):
            load_pc(pc)
        found = set()
        for gp in gp_by_pc.values():
            if not gp:
                continue
            agg, sts, dps = gp
            for act, st_set in sts.items():
                if act.startswith("GLE-") and act_in_exec(st_set, exec_str):
                    found.add(act)
        acts = sorted(found, key=sort_key)
        if not acts:
            return die("Collect: ни одного акта с выполнением %s не найдено." % exec_str)
        last_b = 2 + len(acts)
        xl.ScreenUpdating = False
        try:
            if old_last >= 3:
                ws.Range(ws.Cells(3, 2), ws.Cells(old_last, 2)).ClearContents()
                ws.Range(ws.Cells(3, 2), ws.Cells(old_last, 2)).Interior.ColorIndex = -4142
                ws.Range(ws.Cells(3, 3), ws.Cells(old_last, 21)).ClearContents()
            ws.Range(ws.Cells(3, 2), ws.Cells(last_b, 2)).Value = \
                tuple((a,) for a in acts)
        finally:
            xl.ScreenUpdating = True
    else:
        last_b = old_last
        if last_b < 3:
            return die("В столбце B листа «%s» нет актов со строки 3." % ws.Name)
        raw = ws.Range(ws.Cells(3, 2), ws.Cells(last_b, 2)).Value
        if not isinstance(raw, tuple):
            raw = ((raw,),)
        acts = [txt(r[0]) for r in raw]
        pcs = []
        for act in acts:
            if act:
                pc = pc_of(title_of(act))
                if pc and pc not in pcs:
                    pcs.append(pc)
        for pc in pcs:
            load_pc(pc)

    rows_out = []
    past_flags = []
    nf = no_src = unk = empty_cnt = first_cnt = 0
    for act in acts:
        if not act:
            empty_cnt += 1
            rows_out.append(["-"] * TOTAL_COLS)
            past_flags.append(False)
            continue

        ttl = title_of(act)
        pc = pc_of(ttl)
        base = [ttl, pc, sec_of(act), km_of(act)]

        if not pc:
            unk += 1
            rows_out.append(base + ["—", "—"] + [0.0] * 11 + [0.0, 0.0, 0.0])
            past_flags.append(False)
            continue

        gp = gp_by_pc.get(pc)
        ob = ob_by_pc.get(pc)

        in_gp = gp is not None and act in gp[0]
        in_ob = ob is not None and act in ob

        if not in_gp and not in_ob:
            nf += 1
            rows_out.append(base + ["НЕ НАЙДЕНО", "НЕ НАЙДЕНО"]
                            + [0.0] * 11 + [0.0, 0.0, 0.0])
            past_flags.append(False)
            continue

        # ---------- данные из Сводной ГП (если есть) ----------
        if in_gp:
            agg, sts, dps = gp
            a = agg[act]
            sum_s, ogz, svarka, prot, weld, ogz_area, cnt = a
            st_list = sorted(sts.get(act, ()))
            g = "; ".join(st_list) if st_list else "(без статуса)"
            past = had_past_exec(st_list, exec_str)
            if not past:
                first_cnt += 1
            dp = ", ".join(sorted(dps.get(act, ()))) if dps.get(act) else "—"
        else:
            sum_s = ogz = svarka = prot = weld = ogz_area = cnt = 0.0
            st_list = []
            g = "(нет в Сводной ГП)"
            past = False
            dp = "—"

        # ---------- данные из Общей сводной (если есть) ----------
        if in_ob:
            o_cnt, o_w = ob[act]
        else:
            o_cnt = o_w = 0.0

        vals = [sum_s, o_w, prot, sum_s - prot, ogz, ogz_area, svarka,
                ogz + svarka, weld, float(cnt), o_cnt]

        # ---------- метизы (независимо от ГП) ----------
        met = met_by_pc.get(pc)
        if met and act in met:
            met_g, met_h, met_j = met[act]
            met_vals = [met_j, met_h, met_g]
        else:
            met_vals = [0.0, 0.0, 0.0]

        row = base + [dp, g] + vals + met_vals
        rows_out.append(row)
        past_flags.append(past)

    bad_rows = [(i, len(r)) for i, r in enumerate(rows_out) if len(r) != TOTAL_COLS]
    if bad_rows:
        detail = "; ".join(f"строка {i+3}: {n} вместо {TOTAL_COLS}" for i, n in bad_rows[:5])
        return die(f"Внутренняя ошибка: рваный блок строк.\n{detail}")
    # ---------- запись через COM ----------
    xl.Calculation = XL_MANUAL
    xl.CalculateBeforeSave = False
    xl.ScreenUpdating = False
    try:
        last_a = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        if last_a > last_b and txt(ws.Cells(last_a, 1).Value) == "ИТОГО":
            ws.Range(ws.Cells(last_a, 1), ws.Cells(last_a, LAST_OUTPUT)).ClearContents()

        clear_last = max(last_b, old_last)
        if clear_last >= 3:
            ws.Range(ws.Cells(3, 3), ws.Cells(clear_last, LAST_OUTPUT)).ClearContents()

        if last_b >= 3:
            ws.Range(ws.Cells(3, 3), ws.Cells(last_b, LAST_OUTPUT)).Value = \
                tuple(tuple(r) for r in rows_out)

        # ---------- порядковые номера в столбце A (только для непустых актов) ----------
        if last_b >= 3:
            a_vals = []
            seq = 0
            for act in acts:
                if act:
                    seq += 1
                    a_vals.append((seq,))
                else:
                    a_vals.append((None,))
            ws.Range(ws.Cells(3, 1), ws.Cells(last_b, 1)).Value = tuple(a_vals)

        ws.Cells(2, 1).Value = "ИТОГО"
        for col in range(9, LAST_OUTPUT + 1):
            letter = chr(64 + col)
            ws.Cells(2, col).Formula = "=SUBTOTAL(109,%s3:%s%d)" % (letter, letter, last_b)
        ws.Range(ws.Cells(2, 1), ws.Cells(2, LAST_OUTPUT)).Font.Bold = True

        ws.Range(ws.Cells(3, 9), ws.Cells(last_b, LAST_OUTPUT)).NumberFormatLocal = "# ##0,00"

        for i, past in enumerate(past_flags):
            r = 3 + i
            if past:
                ws.Cells(r, 2).Interior.Color = LIGHT_BLUE
            else:
                ws.Cells(r, 2).Interior.ColorIndex = -4142

        try:
            if ws.AutoFilterMode:
                ws.AutoFilterMode = False
            ws.Range(ws.Cells(2, 1), ws.Cells(last_b, LAST_TABLE)).AutoFilter()
        except Exception:
            pass
        try:
            ws.Activate()
            xl.ActiveWindow.FreezePanes = False
            ws.Range("A3").Select()
            xl.ActiveWindow.FreezePanes = True
        except Exception:
            pass
    finally:
        xl.ScreenUpdating = True

    msg = ("Готово за %.0f с.\nКнига «%s», лист «%s», выполнение %s.\n"
           "Строк: %d (пустых: %d); НЕ НАЙДЕНО: %d; нет источника: %d; ПК не распознан: %d.\n"
           "Впервые (без прошлых выполнений): %d." %
           (time.time() - t0, wbk.Name, ws.Name, exec_str,
            len(acts), empty_cnt, nf, no_src, unk, first_cnt))
    if collect:
        msg += "\nСписок актов пересобран из сводных по всем ПК."
    if warns:
        msg += "\n\nПредупреждения:\n" + "\n".join(warns)
    print(msg)
    win32api.MessageBox(0, msg, "act_analysis_branch_DP", win32con.MB_ICONINFORMATION)
    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        import traceback
        err_path = Path(__file__).resolve().parent / "last_error.txt"
        tb = traceback.format_exc()
        with open(err_path, "w", encoding="utf-8") as f:
            f.write(tb)
        print("КРИТИЧЕСКАЯ ОШИБКА:\n" + tb)
        print("\nОшибка записана в:", err_path)
        try:
            win32api.MessageBox(0, tb, "Анализ актов v" + VERSION, win32con.MB_ICONERROR)
        except Exception:
            pass
        input("Нажмите Enter для закрытия окна...")
        sys.exit(1)