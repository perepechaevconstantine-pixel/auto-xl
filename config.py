# -*- coding: utf-8 -*-
"""
Конфигурация auto-xl: пути к сетевым источникам + система логирования (v3.1.0).

Здесь можно ЖЁСТКО задать шаблоны путей — они имеют приоритет над config.json.
{pc} — номер Пункта Комплексной сборки (1, 2, 3).

Если эти параметры не заданы (или файл удалён), скрипт возьмёт пути из
config.json; если и его нет — спросит папки при первом запуске и сохранит.

CLI-переопределение (высший приоритет):
    python act_analysis.py --gp "D:\\sources\\Сводная_ГП_КМ_ПК{pc}.xlsb" ^
                           --ob "D:\\sources\\ОБЩАЯ...(ПК{pc}).xlsb" [collect]
"""

# GP_PATH_TEMPLATE = (
#     r"\\vls.lan\ULSPG-DFS\ПТО\04. Завод СПГ_Рабочая общая"
#     r"\02. Отчёты по исполнительной документации\Группа КМ\3. Генподряд"
#     r"\!СВОДНЫЕ ГП КМ\Сводная_ГП_КМ_ПК{pc}.xlsb"
# )
# OB_PATH_TEMPLATE = (
#     r"\\vls.lan\ULSPG-DFS\ПТО\04. Завод СПГ_Рабочая общая"
#     r"\02. Отчёты по исполнительной документации"
#     r"\ОБЩАЯ Сводная таблица по статусу объекта СПГ КМ (по линиям)"
#     r"\ОБЩАЯ Сводная таблица по статусу объекта СПГ КМ (ПК{pc}).xlsb"
# )

# ------------------ логирование (v3.1.0) ------------------
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path as _Path

LOG = {
    "level": "INFO",              # DEBUG/INFO/WARNING/ERROR/CRITICAL
    "dir": "logs",                # папка относительно скрипта
    "file": "auto-xl.log",
    "error_file": "errors.log",
    "to_console": True,
    "to_file": True,
    "max_bytes": 5 * 1024 * 1024, # 5 МБ
    "backup_count": 3,
}


def setup_logging(quiet=False):
    """Инициализация логгера: консоль + logs/auto-xl.log + logs/errors.log.

    Вызывается ОДИН раз из act_analysis.main(). Повторные вызовы идемпотентны
    (существующие хендлеры снимаются). quiet=True -> в консоль только ERROR+.
    """
    log_dir = _Path(__file__).resolve().parent / LOG["dir"]
    log_dir.mkdir(parents=True, exist_ok=True)

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)  # всегда DEBUG, фильтрация на хендлерах
    for h in list(root.handlers):
        root.removeHandler(h)
        try:
            h.close()
        except Exception:
            pass

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # 1. Консоль (stderr); при --quiet — только ошибки
    if LOG["to_console"]:
        ch = logging.StreamHandler()
        ch.setLevel(logging.ERROR if quiet else LOG.get("level", "INFO"))
        ch.setFormatter(formatter)
        root.addHandler(ch)

    if LOG["to_file"]:
        # 2. Общий файл: всё до DEBUG
        fh = RotatingFileHandler(
            log_dir / LOG["file"],
            maxBytes=LOG["max_bytes"],
            backupCount=LOG["backup_count"],
            encoding="utf-8",
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        root.addHandler(fh)

        # 3. Файл ошибок: только ERROR и CRITICAL
        eh = RotatingFileHandler(
            log_dir / LOG["error_file"],
            maxBytes=LOG["max_bytes"],
            backupCount=LOG["backup_count"],
            encoding="utf-8",
        )
        eh.setLevel(logging.ERROR)
        eh.setFormatter(formatter)
        root.addHandler(eh)

    return root
