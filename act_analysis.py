# -*- coding: utf-8 -*-
"""
Анализ актов v3.2.0 (основная ветка разработки).
История изменений — см. CHANGELOG.md.
База: v3.1.0 (структурное логирование, --quiet, замеры времени).
Добавлено в v3.2.0 (оптимизация производительности, точность Decimal сохранена):
  - COM-запись данных массивом Range.Value = tuple-of-tuples (было по ячейкам);
  - пакетное форматирование целыми диапазонами (NumberFormat, Font.Bold);
  - отключение Calculation/ScreenUpdating/EnableEvents/DisplayAlerts на всю
    операцию с сохранением/восстановлением исходных значений в finally;
  - кэш прочитанных файлов в памяти на время одного запуска (config.PERF["cache"]);
  - компилированные regex для распознавания ПК (pc_of);
  - openpyxl удалён из зависимостей (не использовался).
Запуск:
    python act_analysis.py                — обновить цифры для текущего списка в B
    python act_analysis.py collect        — пересобрать список актов из сводных
    python act_analysis.py --quiet ...    — тихий режим (в консоль только ошибки)
    python act_analysis.py --gp PATH --ob PATH [collect]  — переопределить пути
"""

import sys
import os
import json
import time
import re
import logging
from functools import lru_cache
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from concurrent.futures import ThreadPoolExecutor

logger = logging.getLogger(__name__)


def _load_perf():
    """Настройки производительности из config.PERF (кэш источников и т.п.)."""
    try:
        import config as _cfg
        return getattr(_cfg, "PERF", {}) or {}
    except Exception:
        return {}


_PERF = _load_perf()
USE_CACHE = bool(_PERF.get("cache", True))
CACHE_TTL = float(_PERF.get("cache_ttl", 300))


def _init_logging(quiet=False):
    """Логгер инициализируется ОДИН раз через config.setup_logging().
    Если config недоступен — минимальный fallback (только консоль), чтобы
    скрипт не падал из-за конфигурации логирования."""
    try:
        import config as _cfg
        if hasattr(_cfg, "setup_logging"):
            return _cfg.setup_logging(quiet=quiet)
    except Exception as e:
        print("Не удалось инициализировать логирование из config.py (%s), "
              "консольный режим по умолчанию" % e)
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    for h in list(root.handlers):
        root.removeHandler(h)
    ch = logging.StreamHandler()
    ch.setLevel(logging.ERROR if quiet else logging.INFO)
    ch.setFormatter(logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"))
    root.addHandler(ch)
    return root

# ------------------ настройки источников (портативные) ------------------
# Источники подключаются шаблоном пути, где {pc} — номер Пункта Комплексной
# сборки (1/2/3). Пути берутся (в порядке приоритета):
#   1. CLI:         --gp PATH --ob PATH
#   2. config.py:   GP_PATH_TEMPLATE / OB_PATH_TEMPLATE (жёсткая настройка)
#   3. config.json: автосоздаётся при первом запуске (диалог выбора папки)
DEFAULT_GP_TEMPLATE = (
    r"\\vls.lan\ULSPG-DFS\ПТО\04. Завод СПГ_Рабочая общая"
    r"\02. Отчёты по исполнительной документации\Группа КМ\3. Генподряд"
    r"\!СВОДНЫЕ ГП КМ\Сводная_ГП_КМ_ПК{pc}.xlsb"
)
DEFAULT_OB_TEMPLATE = (
    r"\\vls.lan\ULSPG-DFS\ПТО\04. Завод СПГ_Рабочая общая"
    r"\02. Отчёты по исполнительной документации"
    r"\ОБЩАЯ Сводная таблица по статусу объекта СПГ КМ (по линиям)"
    r"\ОБЩАЯ Сводная таблица по статусу объекта СПГ КМ (ПК{pc}).xlsb"
)

BASE_DIR = Path(__file__).resolve().parent
CONFIG_JSON = BASE_DIR / "config.json"
LOGS_DIR = BASE_DIR / "logs"


def load_config():
    """-> (gp_template, ob_template). config.py имеет приоритет над config.json."""
    gp, ob = DEFAULT_GP_TEMPLATE, DEFAULT_OB_TEMPLATE
    cfg_py = BASE_DIR / "config.py"
    if cfg_py.exists():
        try:
            import config as _cfg
            # берём только реально заданные значения; закомментированный
            # шаблон не должен отключать config.json
            py_gp = getattr(_cfg, "GP_PATH_TEMPLATE", None)
            py_ob = getattr(_cfg, "OB_PATH_TEMPLATE", None)
            if py_gp or py_ob:
                return py_gp or gp, py_ob or ob
        except Exception as e:
            logger.warning("config.py не загружен (%s), используем config.json", e)
    if CONFIG_JSON.exists():
        try:
            with open(CONFIG_JSON, encoding="utf-8") as f:
                data = json.load(f)
            gp = data.get("gp_template", gp)
            ob = data.get("ob_template", ob)
        except Exception as e:
            logger.warning("config.json повреждён (%s), используем значения по умолчанию", e)
    return gp, ob


def ask_folder(title):
    """Диалог выбора папки (Windows). Fallback — ввод путём в консоли."""
    path = None
    try:
        import win32com.client as _wc
        shell = _wc.Dispatch("Shell.Application")
        folder = shell.BrowseForFolder(0, title, 0x0051, "")
        if folder is not None:
            path = folder.Self.Path
    except Exception:
        pass
    if path is None:
        try:
            path = input(title.replace("\n", " ") + " Введите путь к папке с источниками: ").strip()
        except EOFError:
            path = ""
    return path or None


def ensure_paths_interactive(gp_tmpl, ob_tmpl):
    """Если файлов по шаблонам нет — спросить папки у пользователя и сохранить config.json."""
    changed = False
    for n in ("1", "2", "3"):
        if os.path.exists(gp_tmpl.format(pc=n)):
            break
    else:
        folder = ask_folder("Не найдены файлы «Сводная_ГП_КМ_ПК*.xlsb».\n"
                            "Выберите папку с Сводными ГП:")
        if folder:
            gp_tmpl = os.path.join(folder, "Сводная_ГП_КМ_ПК{pc}.xlsb")
            changed = True
    for n in ("1", "2", "3"):
        if os.path.exists(ob_tmpl.format(pc=n)):
            break
    else:
        folder = ask_folder("Не найдены файлы «ОБЩАЯ Сводная ... (ПК*).xlsb».\n"
                            "Выберите папку с Общей сводной:")
        if folder:
            ob_tmpl = os.path.join(
                folder, "ОБЩАЯ Сводная таблица по статусу объекта СПГ КМ (ПК{pc}).xlsb")
            changed = True
    if changed:
        try:
            with open(CONFIG_JSON, "w", encoding="utf-8") as f:
                json.dump({"gp_template": gp_tmpl, "ob_template": ob_tmpl},
                          f, ensure_ascii=False, indent=2)
            logger.info("Пути сохранены в %s", CONFIG_JSON)
        except Exception as e:
            logger.error("Не удалось сохранить config.json: %s", e)
    return gp_tmpl, ob_tmpl


GP_TEMPLATE, OB_TEMPLATE = load_config()

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

VERSION = "3.2.0"

CENT = Decimal("0.01")

# v3.2.0: компилированные regex для распознавания ПК (было: срезы строк)
PC1_RE = re.compile(r"^L03|^L3-02-010\.|^L3-12-030\.")
PC2_RE = re.compile(r"^L0?1(?!\d)")   # L1xxx / L01xxx, но не L10/L12 (это ПК3)
PC3_RE = re.compile(r"^L3")

# ------------------ вспомогательные функции ------------------
def dec2(v):
    """Приводит любое значение (int, float, str, Decimal, None) к Decimal
    с арифметическим округлением до 2 знаков (ROUND_HALF_UP).
    Нечисловые значения трактуются как 0.00. Фикс бага v2.9.9:
    'int' object has no attribute 'quantize'."""
    if v is None or isinstance(v, bool):
        return Decimal("0.00")
    if isinstance(v, Decimal):
        d = v
    elif isinstance(v, int):
        d = Decimal(v)
    elif isinstance(v, float):
        d = Decimal(repr(v))
    else:
        s = txt(v).replace(" ", "").replace(",", ".")
        try:
            d = Decimal(s)
        except Exception:
            return Decimal("0.00")
    return d.quantize(CENT, rounding=ROUND_HALF_UP)

def num_dec(v):
    """Как num(), но возвращает Decimal (или None, если значения нет)."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, Decimal):
        return v
    if isinstance(v, (int, float)):
        return dec2(v)
    s = txt(v).replace(" ", "").replace(",", ".")
    if not s:
        return None
    try:
        return dec2(s)
    except Exception:
        return None

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
    """v3.2.0: компилированные regex вместо срезов строк."""
    if PC1_RE.search(ttl):
        return "ПК1"
    if PC2_RE.match(ttl):
        return "ПК2"
    if PC3_RE.match(ttl):
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
_ROWS_CACHE = {}   # key -> (rows, timestamp); живёт только в памяти процесса


def _rows_cached(path, sheet_name, ttl=None):
    """v3.2.0: кэш строк листа по (путь, лист, mtime, размер файла).

    Внутри одного запуска повторное чтение того же файла не ходит в сеть
    (например, ГП-файл читают read_gp и read_metizy). Кэш — в памяти
    процесса, при заверении скрипта сбрасывается сам. Инвалидация — по
    mtime/размеру; TTL — страховка от «mtime не изменился, содержимое
    другое» на некоторых сетевых шарах. Управляется config.PERF.
    """
    if ttl is None:
        ttl = CACHE_TTL
    from python_calamine import CalamineWorkbook
    if not USE_CACHE:
        wb = CalamineWorkbook.from_path(str(path))
        return wb.get_sheet_by_name(sheet_name).to_python()
    try:
        st = os.stat(str(path))
        key = (str(path), sheet_name, int(st.st_mtime), st.st_size)
    except OSError:
        key = None
    now = time.time()
    if key is not None:
        hit = _ROWS_CACHE.get(key)
        if hit and now - hit[1] < ttl:
            logger.debug("Кэш HIT: %s [%s]", path, sheet_name)
            return hit[0]
    wb = CalamineWorkbook.from_path(str(path))
    rows = wb.get_sheet_by_name(sheet_name).to_python()
    if key is not None:
        _ROWS_CACHE[key] = (rows, now)
        stale = [k for k, (_, ts) in _ROWS_CACHE.items() if now - ts >= ttl]
        for k in stale:
            _ROWS_CACHE.pop(k, None)
    return rows


def read_gp(path, exec_str):
    """-> (agg: act->[sumS, ogz, svarka, prot, weld, ogz_area, cnt],
           sts: act->set, dps: act->set)
    Все суммы и счётчики — Decimal (точные денежные расчёты)."""
    rows = _rows_cached(path, "ГПВ КМ")
    agg, sts, dps = {}, {}, {}
    Z = Decimal("0.00")
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
        s = num_dec(cell(row, GP_S))
        w = num_dec(cell(row, GP_W))
        area = num_dec(cell(row, GP_AB))
        a = agg.get(act)
        if a is None:
            a = [Z, Z, Z, Z, Z, Z, Decimal(0)]
            agg[act] = a
        if s is not None:
            a[0] += s
        cls = classify(st, exec_str)
        if cls == "ogz":
            a[1] += s if s else Z
            a[5] += area if area else Z
        elif cls == "svarka":
            a[2] += s if s else Z
        elif cls == "prot":
            a[3] += s if s else Z
        if w is not None:
            a[4] += w
        a[6] += 1
        if st:
            sts.setdefault(act, set()).add(st)
    return agg, sts, dps

def read_ob(path):
    """-> agg: act->[cnt, sumW]. cnt/sumW — Decimal."""
    rows = _rows_cached(path, "Сводная КМ")
    agg = {}
    for row in rows[SKIP_ROWS:]:
        act = txt(cell(row, OB_AO))
        if not act:
            continue
        if not txt(cell(row, OB_J)):
            continue
        w = num_dec(cell(row, OB_W))
        a = agg.get(act)
        if a is None:
            a = [Decimal(0), Decimal("0.00")]
            agg[act] = a
        a[0] += 1
        if w is not None:
            a[1] += w
    return agg

def read_metizy(path):
    """-> agg: act->[sum_G, sum_H, sum_J] (к освоению, ВПБ, по ИД). Decimal."""
    rows = _rows_cached(path, "Метизы")
    agg = {}
    for row in rows[MET_SKIP_ROWS:]:
        act = txt(cell(row, MET_D))
        if not act:
            continue
        g = num_dec(cell(row, MET_G))
        h = num_dec(cell(row, MET_H))
        j = num_dec(cell(row, MET_J))
        a = agg.get(act)
        if a is None:
            a = [Decimal("0.00"), Decimal("0.00"), Decimal("0.00")]
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
    print("Анализ актов v" + VERSION)  # первая строка консоли (добавлено в v3.0.2)

    quiet = "--quiet" in sys.argv[1:]
    _init_logging(quiet=quiet)
    logger.info("Запуск с аргументами: %s", sys.argv[1:])
    try:
        import config as _cfg
        logger.debug("Конфигурация LOG: %s", getattr(_cfg, "LOG", None))
    except Exception:
        pass
    logger.debug("GP_TEMPLATE=%s | OB_TEMPLATE=%s", GP_TEMPLATE, OB_TEMPLATE)

    import win32api
    import win32con
    import win32com.client as wc

    t0 = time.time()

    def die(msg, code=1):
        logger.error(msg.replace("\n", " | "))
        win32api.MessageBox(0, msg, "Анализ актов v" + VERSION, win32con.MB_ICONERROR)
        print(msg)
        return code

    collect = False
    gp_tmpl, ob_tmpl = GP_TEMPLATE, OB_TEMPLATE
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        a = args[i].strip().lower()
        if a == "--gp" and i + 1 < len(args):
            gp_tmpl = args[i + 1]
            i += 2
        elif a == "--ob" and i + 1 < len(args):
            ob_tmpl = args[i + 1]
            i += 2
        elif a == "collect":
            collect = True
            i += 1
        else:
            i += 1
    logger.info("Режим: %s; ГП-шаблон: %s; ОБЩ-шаблон: %s",
                "collect" if collect else "update", gp_tmpl, ob_tmpl)

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


    # v3.2.0: отключаем автооперации Excel НА ВСЮ операцию сразу после
    # получения xlApp; исходные значения сохраняем и восстанавливаем в
    # finally (включая collect-режим и ранние выходы через die()).
    old_calc = old_events = old_alerts = None
    try:
        old_calc = xl.Calculation
    except Exception:
        pass
    try:
        old_events = xl.EnableEvents
    except Exception:
        old_events = True
    try:
        old_alerts = xl.DisplayAlerts
    except Exception:
        old_alerts = True
    try:
        xl.Calculation = XL_MANUAL
    except Exception:
        pass
    try:
        xl.CalculateBeforeSave = False
    except Exception:
        pass
    try:
        xl.EnableEvents = False
    except Exception:
        pass
    try:
        xl.DisplayAlerts = False
    except Exception:
        pass
    xl.ScreenUpdating = False
    logger.debug("Excel: Calculation/ScreenUpdating/EnableEvents/DisplayAlerts отключены")

    try:  # v3.2.0: внешний try — finally ниже всегда восстановит Excel
      old_last = ws.Cells(ws.Rows.Count, 2).End(XL_UP).Row

      gp_by_pc, ob_by_pc, met_by_pc = {}, {}, {}
      warns = []

      # v3.0.1: если пути из конфига не ведут к файлам — спросим у пользователя
      if not any(os.path.exists(gp_tmpl.format(pc=n)) for n in ("1", "2", "3")) or \
         not any(os.path.exists(ob_tmpl.format(pc=n)) for n in ("1", "2", "3")):
          gp_tmpl, ob_tmpl = ensure_paths_interactive(gp_tmpl, ob_tmpl)

      def load_pc(pc):
          n = pc[2]
          t1 = time.time()
          gp_path = gp_tmpl.format(pc=n)
          try:
              logger.debug("Чтение ГП: %s", gp_path)
              gp_by_pc[pc] = read_gp(gp_path, exec_str)
          except Exception as e:
              gp_by_pc[pc] = None
              warns.append("Нет источника ГП по %s (%s)." % (pc, e))
              logger.warning("Нет источника ГП по %s", pc, exc_info=True)
          try:
              ob_path = ob_tmpl.format(pc=n)
              logger.debug("Чтение Общей сводной: %s", ob_path)
              ob_by_pc[pc] = read_ob(ob_path)
          except Exception as e:
              ob_by_pc[pc] = None
              warns.append("Нет источника Общей по %s (%s)." % (pc, e))
              logger.warning("Нет источника Общей по %s", pc, exc_info=True)
          try:
              met_by_pc[pc] = read_metizy(gp_path)
          except Exception as e:
              met_by_pc[pc] = None
              warns.append("Нет листа Метизы по %s (%s)." % (pc, e))
              logger.warning("Нет листа Метизы по %s", pc, exc_info=True)
          elapsed = time.time() - t1
          logger.info("%s: чтение источников завершено за %.1f с", pc, elapsed)

      def load_all(pcs):
          """v3.0.0: параллельное чтение источников по ПК (I/O-bound, сеть)."""
          t_load = time.time()
          logger.info("Загрузка файлов: параллельное чтение %d ПК", len(pcs))
          with ThreadPoolExecutor(max_workers=max(1, len(pcs))) as ex:
              list(ex.map(load_pc, pcs))
          logger.info("Загрузка файлов завершена за %.2f с", time.time() - t_load)

      if collect:
          load_all(("ПК1", "ПК2", "ПК3"))
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
          # v3.2.0: ScreenUpdating уже выключен на всю операцию (см. начало main)
          if old_last >= 3:
              ws.Range(ws.Cells(3, 2), ws.Cells(old_last, 2)).ClearContents()
              ws.Range(ws.Cells(3, 2), ws.Cells(old_last, 2)).Interior.ColorIndex = -4142
              ws.Range(ws.Cells(3, 3), ws.Cells(old_last, 21)).ClearContents()
          ws.Range(ws.Cells(3, 2), ws.Cells(last_b, 2)).Value = \
              tuple((a,) for a in acts)
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
          load_all(pcs)

      t_parse = time.time()
      Z = Decimal("0.00")
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
              rows_out.append(base + ["—", "—"] + [Z] * 14)
              past_flags.append(False)
              continue

          gp = gp_by_pc.get(pc)
          ob = ob_by_pc.get(pc)

          in_gp = gp is not None and act in gp[0]
          in_ob = ob is not None and act in ob

          if not in_gp and not in_ob:
              nf += 1
              rows_out.append(base + ["НЕ НАЙДЕНО", "НЕ НАЙДЕНО"] + [Z] * 14)
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
              sum_s = ogz = svarka = prot = weld = ogz_area = cnt = Z
              st_list = []
              g = "(нет в Сводной ГП)"
              past = False
              dp = "—"

          # ---------- данные из Общей сводной (если есть) ----------
          if in_ob:
              o_cnt, o_w = ob[act]
          else:
              o_cnt = o_w = Z

          # точные денежные расчёты в Decimal с квантованием до 2 знаков
          vals = [dec2(sum_s), dec2(o_w), dec2(prot), dec2(sum_s - prot),
                  dec2(ogz), dec2(ogz_area), dec2(svarka),
                  dec2(ogz + svarka), dec2(weld), dec2(cnt), dec2(o_cnt)]

          # ---------- метизы (независимо от ГП) ----------
          met = met_by_pc.get(pc)
          if met and act in met:
              met_g, met_h, met_j = met[act]
              met_vals = [dec2(met_j), dec2(met_h), dec2(met_g)]
          else:
              met_vals = [Z, Z, Z]

          row = base + [dp, g] + vals + met_vals
          rows_out.append(row)
          past_flags.append(past)

      bad_rows = [(i, len(r)) for i, r in enumerate(rows_out) if len(r) != TOTAL_COLS]
      logger.info("Обработка/парсинг данных завершена за %.2f с (%d строк)",
                  time.time() - t_parse, len(acts))
      if bad_rows:
          detail = "; ".join(f"строка {i+3}: {n} вместо {TOTAL_COLS}" for i, n in bad_rows[:5])
          return die(f"Внутренняя ошибка: рваный блок строк.\n{detail}")

      # ---------- запись через COM ----------
      t_write = time.time()
      # v3.2.0: Excel-настройки уже отключены выше; данные пишутся массивом
      # одним вызовом Range.Value; восстановление — во внешнем finally.
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

              # ---------- пакетная подсветка столбца B (v3.0.0) ----------
              # Красим блоками одного цвета: прогоняем past_flags и выделяем
              # непрерывные серии одинакового состояния — меньше вызовов COM.
              if last_b >= 3:
                  n = len(past_flags)
                  i = 0
                  while i < n:
                      j = i
                      while j < n and past_flags[j] == past_flags[i]:
                          j += 1
                      r1, r2 = 3 + i, 3 + j - 1
                      rng = ws.Range(ws.Cells(r1, 2), ws.Cells(r2, 2))
                      if past_flags[i]:
                          rng.Interior.Color = LIGHT_BLUE
                      else:
                          rng.Interior.ColorIndex = -4142
                      i = j

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
        try:
            xl.EnableEvents = old_events
        except Exception:
            pass
        try:
            xl.DisplayAlerts = old_alerts
        except Exception:
            pass
        if old_calc is not None:
            try:
                xl.Calculation = old_calc
            except Exception:
                pass

    msg = ("Готово за %.0f с.\nКнига «%s», лист «%s», выполнение %s.\n"
           "Строк: %d (пустых: %d); НЕ НАЙДЕНО: %d; нет источника: %d; ПК не распознан: %d.\n"
           "Впервые (без прошлых выполнений): %d." %
           (time.time() - t0, wbk.Name, ws.Name, exec_str,
            len(acts), empty_cnt, nf, no_src, unk, first_cnt))
    if collect:
        msg += "\nСписок актов пересобран из сводных по всем ПК."
    if warns:
        msg += "\n\nПредупреждения:\n" + "\n".join(warns)
    logger.info("Запись результата в Excel завершена за %.2f с", time.time() - t_write)
    logger.info("Итог: строк=%d пустых=%d НЕ_НАЙДЕНО=%d нет_источника=%d "
                "ПК_не_распознан=%d впервые=%d предупреждений=%d",
                len(acts), empty_cnt, nf, no_src, unk, first_cnt, len(warns))
    print(msg)  # финальный результат — всегда в консоль, даже в --quiet
    win32api.MessageBox(0, msg, "Анализ актов v" + VERSION, win32con.MB_ICONINFORMATION)
    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception as exc:
        # логгер мог ещё не инициализироваться (ошибка до main()) — инициируем
        _init_logging(quiet="--quiet" in sys.argv[1:])
        logger.critical("КРИТИЧЕСКАЯ ОШИБКА", exc_info=True)
        tb = __import__("traceback").format_exc()
        here = Path(__file__).resolve().parent
        with open(here / "last_error.txt", "w", encoding="utf-8") as f:
            f.write(tb)
        print("КРИТИЧЕСКАЯ ОШИБКА:\n" + tb)
        print("\nОшибка записана в:", here / "last_error.txt",
              "и в логи (logs/auto-xl.log, logs/errors.log)")
        try:
            import win32api, win32con
            win32api.MessageBox(0, tb, "Анализ актов v" + VERSION, win32con.MB_ICONERROR)
        except Exception:
            pass
        input("Нажмите Enter для закрытия окна...")
        sys.exit(1)