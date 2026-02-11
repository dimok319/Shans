# -*- coding: utf-8 -*-
import os
import re
import shutil
from datetime import datetime, timedelta
from pathlib import Path

import pythoncom
import win32com.client as win32

# ----------------------------
# НАСТРОЙКИ
# ----------------------------
SRC_DIR = r"\\192.168.1.211\аналитический центр\Отчёты\ДЗ\Контроль сроков оплаты"
DEST_DIR = r"\\192.168.1.211\аналитический центр\Пышнограев"

DATA_FILE = r"\\192.168.1.211\дебиторская задолженность\Отчеты\Выгрузка 1С\ДЗ_1С_НОВЫЙ (XLSX).xlsx"
SHEET_NAME = "TDSheet"

TARGET_HEADER_ROWS_TO_KEEP = 1  # Шапка: оставляем 2 строки
HEADER_SCAN_COLS = 300
HEADER_SCAN_ROWS = 120

EXCEL_VISIBLE = False  # Включите, если хотите видеть Excel
LOG_PATH = Path(__file__).with_name("dz_update_log.txt")


# ----------------------------
# ЛОГИ
# ----------------------------
def log(msg: str):
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S} | {msg}"
    print(line)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ----------------------------
# ВЫБОР ФАЙЛА-ИСТОЧНИКА (<= ВЧЕРА)
# ----------------------------
def pick_report_as_of_yesterday(src_dir: str) -> str:
    p = Path(src_dir)
    files = list(p.glob("Отчет по ДЗ*.xlsx")) + list(p.glob("Отчёт по ДЗ*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]

    if not files:
        raise FileNotFoundError(f"В папке нет файлов 'Отчет по ДЗ*.xlsx' / 'Отчёт по ДЗ*.xlsx': {src_dir}")

    cutoff = (datetime.now() - timedelta(days=1)).date()
    best_file, best_date = None, None

    for f in files:
        m = re.search(r"(\d{2}\.\d{2}\.\d{4})", f.stem)
        if not m:
            continue
        try:
            d = datetime.strptime(m.group(1), "%d.%m.%Y").date()
        except ValueError:
            continue

        if d <= cutoff and (best_date is None or d > best_date):
            best_date, best_file = d, f

    if best_file is None:
        raise FileNotFoundError(f"Не найден отчёт с датой в имени <= вчера ({cutoff:%d.%m.%Y}).")

    return str(best_file)


def extract_date_from_filename(path: str) -> str:
    name = Path(path).stem
    m = re.search(r"(\d{2}\.\d{2}\.\d{4})", name)
    return m.group(1) if m else (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")


def make_unique_dest_path(dest_dir: str, base_filename: str) -> str:
    dest_dir_p = Path(dest_dir)
    candidate = dest_dir_p / base_filename
    if not candidate.exists():
        return str(candidate)

    ts = datetime.now().strftime("%H%M%S")
    c2 = dest_dir_p / f"{candidate.stem}_{ts}{candidate.suffix}"
    if not c2.exists():
        return str(c2)

    for i in range(1, 100):
        c = dest_dir_p / f"{candidate.stem}_run{i}{candidate.suffix}"
        if not c.exists():
            return str(c)

    raise RuntimeError("Не удалось подобрать свободное имя файла в DEST_DIR.")


# ----------------------------
# Excel helpers
# ----------------------------
def normalize_header(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def build_alias_map():
    return {
        "Клиент.Новый клиент сезона": "Новый клиент сезона",
        "Новый клиент сезона": "Новый клиент сезона",
        "Контрагент.Сокращенное наименование": "Контрагент.Сокращенное юр. наименование",
        "Контрагент.Сокращенное юр. наименование": "Контрагент.Сокращенное юр. наименование",
        "Общая дебиторская задолженность,руб": "Общая дебиторская задолженность, руб",
        "Общая дебиторская задолженность, руб": "Общая дебиторская задолженность, руб",
    }


def remap_header(h: str, alias_map: dict) -> str:
    h0 = normalize_header(h)
    return alias_map.get(h0, h0)


def find_header_row(ws, must_have=("Дивизион", "Регион.Наименование"),
                    scan_rows=HEADER_SCAN_ROWS, scan_cols=HEADER_SCAN_COLS):
    must_have = [normalize_header(x) for x in must_have]
    used_rows = int(ws.UsedRange.Rows.Count) if ws.UsedRange is not None else scan_rows
    max_row = min(used_rows if used_rows > 0 else scan_rows, scan_rows)

    for r in range(1, max_row + 1):
        vals = ws.Range(ws.Cells(r, 1), ws.Cells(r, scan_cols)).Value
        row_vals = list(vals[0]) if isinstance(vals, tuple) and len(vals) > 0 else []
        row_norm = [normalize_header(v) for v in row_vals]
        row_set = set([v for v in row_norm if v])

        if row_set and all(x in row_set for x in must_have):
            last_col = scan_cols
            while last_col > 1 and (row_norm[last_col - 1] == "" or row_norm[last_col - 1] is None):
                last_col -= 1
            return r, row_norm[:last_col], last_col

    raise RuntimeError(f"Не удалось найти строку заголовков на листе '{ws.Name}'")


def get_last_data_row(ws, key_col: int, start_row: int) -> int:
    # xlUp = -4162
    last = int(ws.Cells(ws.Rows.Count, key_col).End(-4162).Row)
    return max(last, start_row)


# ----------------------------
# MAIN
# ----------------------------
def main():
    if LOG_PATH.exists():
        LOG_PATH.unlink()

    # 1) выбираем исходник (последняя прошедшая дата <= вчера)
    src_path = pick_report_as_of_yesterday(SRC_DIR)
    date_str = extract_date_from_filename(src_path)

    os.makedirs(DEST_DIR, exist_ok=True)
    dest_path = make_unique_dest_path(DEST_DIR, f"Отчет по ДЗ {date_str}.xlsx")

    log(f"Выбран исходный файл: {src_path}")
    log(f"Создаю копию: {dest_path}")
    shutil.copy2(src_path, dest_path)
    log(f"Копия создана. Размер: {os.path.getsize(dest_path):,} bytes")

    excel = None
    wb_target = None
    wb_data = None

    pythoncom.CoInitialize()
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = EXCEL_VISIBLE
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        # вручную пересчитываем, чтобы ускорить
        try:
            excel.Calculation = -4135  # xlManual
        except Exception:
            pass

        log("Открываю целевой файл (копию)...")
        wb_target = excel.Workbooks.Open(dest_path, UpdateLinks=0, ReadOnly=False)
        ws_t = wb_target.Worksheets(SHEET_NAME)

        log("Открываю источник данных (считать и закрыть)...")
        wb_data = excel.Workbooks.Open(DATA_FILE, UpdateLinks=0, ReadOnly=True)
        ws_s = wb_data.Worksheets(SHEET_NAME)

        alias = build_alias_map()

        src_header_row, src_headers, _ = find_header_row(ws_s)
        tgt_header_row, tgt_headers, tgt_last_col = find_header_row(ws_t)

        log(f"Заголовок источника: строка {src_header_row}")
        log(f"Заголовок цели: строка {tgt_header_row}, колонок {tgt_last_col}")

        src_map = {}
        for i, h in enumerate(src_headers, start=1):
            hh = remap_header(h, alias)
            if hh and hh not in src_map:
                src_map[hh] = i

        tgt_map = {}
        for i, h in enumerate(tgt_headers, start=1):
            hh = remap_header(h, alias)
            if hh and hh not in tgt_map:
                tgt_map[hh] = i

        common = [h for h in tgt_map.keys() if h in src_map]
        log(f"Сопоставлено общих столбцов: {len(common)}")
        if not common:
            raise RuntimeError("Нет совпадающих заголовков между источником и целью.")

        data_start_t = max(tgt_header_row + 1, TARGET_HEADER_ROWS_TO_KEEP + 1)
        data_start_s = src_header_row + 1

        key_candidates = ["Договор", "Заказ клиента", "Дивизион"]
        key_col_s = next((src_map[k] for k in key_candidates if k in src_map), 1)
        last_row_s = get_last_data_row(ws_s, key_col_s, data_start_s)
        rows_count = max(0, last_row_s - data_start_s + 1)
        log(f"Источник: {rows_count} строк данных ({data_start_s}..{last_row_s})")
        if rows_count == 0:
            raise RuntimeError("В источнике не найдено строк данных.")

        # очистка цели
        key_col_t = next((tgt_map[k] for k in key_candidates if k in tgt_map), 1)
        last_row_t = get_last_data_row(ws_t, key_col_t, data_start_t)
        log(f"Цель: очищаю старые данные {data_start_t}..{last_row_t}")
        ws_t.Range(ws_t.Cells(data_start_t, 1), ws_t.Cells(max(last_row_t, data_start_t), tgt_last_col)).ClearContents()

        # сбор массива на вставку
        log("Собираю массив для вставки...")
        out = [[None] * tgt_last_col for _ in range(rows_count)]
        for h in common:
            s_col = src_map[h]
            t_col = tgt_map[h]
            col_vals = ws_s.Range(ws_s.Cells(data_start_s, s_col), ws_s.Cells(last_row_s, s_col)).Value
            for i in range(rows_count):
                out[i][t_col - 1] = col_vals[i][0] if isinstance(col_vals[i], tuple) else col_vals[i]

        log("Вставляю данные в цель одним диапазоном...")
        ws_t.Range(ws_t.Cells(data_start_t, 1), ws_t.Cells(data_start_t + rows_count - 1, tgt_last_col)).Value = tuple(tuple(r) for r in out)
        log("Вставка завершена.")

        # Закрываем источник
        wb_data.Close(SaveChanges=False)
        wb_data = None
        log("Источник закрыт.")

        log("Сохраняю целевой файл...")
        wb_target.Save()

        log("ГОТОВО ✅")
        log(f"Файл: {dest_path}")

    finally:
        try:
            if wb_data is not None:
                wb_data.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if wb_target is not None:
                wb_target.Close(SaveChanges=True)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
        log(f"ЛОГ: {LOG_PATH}")


if __name__ == "__main__":
    main()
