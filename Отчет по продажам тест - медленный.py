import os
import glob
import time
import shutil
import threading
from datetime import datetime

import pandas as pd
import xlwings as xw

import colorama
from colorama import Fore, Style
colorama.init(autoreset=True)

import win32com.client as win32
import pythoncom


# --------------------------- Настройки адресатов ---------------------------

division_email_map = {
    "Дивизион ЦЕНТР": {
        "to": ["yuchasovskih@shans-group.com", "stepin@shans-group.com", "apanin@shans-group.com",
               "a.smalev@shans-group.com", "aemelyanov@shans-group.com", "vbukreev@shans-group.com",
               "akuryakov@shans-group.com", "a.sadzhaya@shans-group.com", "estrelnikov@shans-group.com",
               "vburykin@shans-group.com"],
        "cc": ["ozakharchenko@shans-group.com", "grigorova@shans-group.com", "rsakhnenko@shans-group.com"],
        "name": "Центр"
    },
    "Дивизион ПОВОЛЖЬЕ": {
        "to": ["ryagafarov@shans-group.com", "ogusev@shans-group.com", "mdadonov@shans-group.com",
               "asemin@shans-group.com", "vkiyaykin@shans-group.com", "acherchimcev@shans-group.com",
               "dkomissarenko@shans-group.com", "oeremin@shans-group.com", "vshuverов@shans-group.com"],
        "cc": ["dsitnikov@shans-group.com", "mcherniy@shans-group.com"],
        "name": "Поволжье"
    },
    "Дивизион ДАЛЬНИЙ ВОСТОК": {
        "to": ["dsolovyev@shans-group.com"],
        "cc": ["pyablonskiy@shans-group.com", "ivorokosov@shans-group.com", "vbashkatov@shans-group.com"],
        "name": "Дальний восток"
    },
    "Дивизион УРАЛ": {
        "to": ["dsolovyev@shans-group.com"],
        "cc": ["pyablonskiy@shans-group.com", "ivorokosov@shans-group.com", "vbashkatov@shans-group.com"],
        "name": "Урал"
    },
    "Дивизион СИБИРЬ": {
        "to": ["dsolovyev@shans-group.com", "gzhukov@shans-group.com", "ichernov@shans-group.com",
               "vvakulin@shans-group.com", "rrihsiev@shans-group.com", "bdolchanin@shans-group.com",
               "epruss@shans-group.com", "s.dudnickiy@shans-group.com", "nivanova@shans-group.com",
               "anikitin@shans-group.com", "mmilkina@shans-group.com"],
        "cc": ["pyablonskiy@shans-group.com", "ivorokosov@shans-group.com", "vbashkatov@shans-group.com"],
        "name": "Сибирь"
    },
    "Дивизион ЮГ": {
        "to": ["ltkachenya@shans-group.com"],
        "cc": [],
        "name": "ЮГ"
    },
}


# --------------------------- Утилиты ---------------------------

def remove_reports_from_subfolders(start_dir):
    """Удаляет временные выгрузки в текущей папке по сегодняшней дате."""
    date_prefix = datetime.today().strftime('%m-%d')
    for dirpath, _, filenames in os.walk(start_dir):
        for filename in filenames:
            if filename.startswith(f"Отчет по продажам {date_prefix}") and filename.endswith(".xlsx"):
                try:
                    os.remove(os.path.join(dirpath, filename))
                except Exception as e:
                    print(f"Ошибка при удалении файла {filename}: {e}")


def safe_clear_sheet(sheet, attempts=5, delay=0.5):
    """Осторожно очищает лист (используем только для рабочих листов с данными, НЕ для листа планов)."""
    import time as _time
    for _ in range(attempts):
        try:
            sheet.api.UsedRange.ClearContents()
            return
        except Exception:
            _time.sleep(delay)
    print(f"Не удалось очистить лист {sheet.name} после {attempts} попыток.")


def send_division_email(division_name):
    pythoncom.CoInitialize()
    try:
        if division_name not in division_email_map:
            print(f"Email для дивизиона '{division_name}' не задан.")
            return
        info = division_email_map[division_name]
        to_list = info["to"]
        cc_list = info["cc"]
        division_display = info["name"]

        body = f"""Коллеги, добрый день!

Отчёт по продажам выложен в папки регионов по пути: \\\\192.168.1.211\\торговый дом\\"Ваш Дивизион"\\"Необходимый регион"\\Маркетинг



-- 
С уважением,
Никита Игоревич Сыроижко
Заместитель Руководителя отдела логистики и отгрузки
394006, г. Воронеж, ул. Ворошилова, д. 1А.
8(473)220-49-41 (доб. 120)
nsyroizhko@shans-group.com
ООО «ШАНС ТРЕЙД» 
Это электронное сообщение и любые документы, приложенные к нему, содержат конфиденциальную информацию. Настоящим уведомляем Вас о том, что если это сообщение не предназначено Вам, использование, копирование, распространение информации, содержащейся в настоящем сообщении, а также осуществление любых действий на основе этой информации, строго запрещено. Если Вы получили это сообщение по ошибке, пожалуйста, сообщите об этом отправителю по электронной почте и удалите это сообщение.
"""

        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = ";".join(to_list)
        mail.CC = ";".join(cc_list)
        mail.Subject = "Отчёт по продажам"
        mail.Body = body
        mail.SentOnBehalfOfName = "Сыроижко Никита Игоревич <nsyroizhko@shans-group.com>"
        mail.Send()
        print(f"Письмо отправлено дивизиону {division_display}")
    except Exception as e:
        print(f"Ошибка отправки письма дивизиону {division_display}: {e}")
    finally:
        pythoncom.CoUninitialize()


# --------------------------- Работа с листом «План контрактации» ---------------------------

REGION_COL_CANDIDATES = ("ОП", "Регион", "Наименование")  # возможные названия столбца региона

def _find_region_col(headers):
    low = [str(h).strip().lower() for h in headers]
    for cand in REGION_COL_CANDIDATES:
        c = cand.strip().lower()
        if c in low:
            return headers[low.index(c)]
    return None


def snapshot_plan_tables(sh_plan):
    """
    Снимаем слепок ВСЕХ таблиц (ListObject) на листе:
    {table_name: {"headers":[...], "df":DataFrame, "anchor_rc":(row,col)}}
    """
    snap = {}
    los = sh_plan.api.ListObjects
    for i in range(1, los.Count + 1):
        lo = los.Item(i)
        name = lo.Name
        headers = [lo.ListColumns(j).Name for j in range(1, lo.ListColumns.Count + 1)]

        if lo.DataBodyRange is None:
            data_vals = []
        else:
            data_vals = sh_plan.range(lo.DataBodyRange.Address).value
            if data_vals is None:
                data_vals = []
            elif data_vals and not isinstance(data_vals[0], (list, tuple)):
                data_vals = [data_vals]

        df_all = pd.DataFrame(data_vals, columns=headers) if data_vals else pd.DataFrame(columns=headers)
        snap[name] = {
            "headers": headers,
            "df": df_all,
            "anchor_rc": (lo.HeaderRowRange.Row, lo.HeaderRowRange.Column),
        }
    return snap


def _resize_and_write_table(sh_plan, lo, headers, df_out, anchor_rc):
    """
    Меняем размер таблицы и записываем значения.
    Плюс: чистим «хвост» старых данных под таблицей, чтобы не оставались строки за пределами ListObject.
    """
    # 1) границы старого тела (до Resize)
    prev_body = lo.DataBodyRange
    prev_first_col = prev_last_row = prev_last_col = None
    if prev_body is not None:
        prev_first_col = prev_body.Column
        prev_last_row  = prev_body.Row + prev_body.Rows.Count - 1
        prev_last_col  = prev_body.Column + prev_body.Columns.Count - 1

    # 2) новый размер
    nrows = max(1, len(df_out))  # число строк тела (без строки заголовка)
    ncols = len(headers)
    hr, hc = anchor_rc

    # Диапазон: заголовок + тело (nrows строк)
    new_range = sh_plan.api.Range(
        sh_plan.api.Cells(hr, hc),
        sh_plan.api.Cells(hr + nrows, hc + ncols - 1)
    )
    lo.Resize(new_range)

    # 3) пишем новое тело
    body_rng = sh_plan.range((hr + 1, hc)).resize(nrows, ncols)
    if len(df_out) == 0:
        body_rng.value = [[""] * ncols]
    else:
        body_rng.value = df_out.values.tolist()

    # 4) чистим «хвост» ниже новой таблицы (в пределах старой ширины)
    if prev_last_row and prev_first_col:
        new_last_row = hr + nrows  # последняя строка новой области тела
        if new_last_row < prev_last_row:
            sh_plan.api.Range(
                sh_plan.api.Cells(new_last_row + 1, prev_first_col),
                sh_plan.api.Cells(prev_last_row, prev_last_col)
            ).ClearContents()


def clean_tails_under_all_tables(sh_plan):
    """
    Доп. страховка: очищаем все данные под каждой таблицей до следующего заголовка таблицы
    (не задевая соседние таблицы).
    """
    los = sh_plan.api.ListObjects
    if los.Count == 0:
        return

    # соберём (name, header_row, first_col, last_col, body_last_row)
    meta = []
    for i in range(1, los.Count + 1):
        lo = los.Item(i)
        hr = lo.HeaderRowRange.Row
        fc = lo.Range.Column
        lc = lo.Range.Column + lo.Range.Columns.Count - 1
        if lo.DataBodyRange is None:
            blr = hr  # тело пустое — ниже заголовка
        else:
            blr = lo.DataBodyRange.Row + lo.DataBodyRange.Rows.Count - 1
        meta.append((lo.Name, hr, fc, lc, blr))

    # сортируем по строке заголовка
    meta.sort(key=lambda t: t[1])

    # граница листа
    used = sh_plan.api.UsedRange
    used_last_row = used.Row + used.Rows.Count - 1

    # чистим под каждой таблицей до следующего заголовка
    for idx, (name, hr, fc, lc, blr) in enumerate(meta):
        # следующая таблица
        next_hr = meta[idx + 1][1] if idx + 1 < len(meta) else used_last_row + 1
        clear_from = blr + 1
        clear_to = min(next_hr - 1, used_last_row)
        if clear_from <= clear_to:
            sh_plan.api.Range(
                sh_plan.api.Cells(clear_from, fc),
                sh_plan.api.Cells(clear_to, lc)
            ).ClearContents()


def filter_plan_sheet_by_region_from_snapshot(sh_plan, snapshot, region_name):
    """
    Для КАЖДОЙ таблицы из snapshot оставляет строки ТОЛЬКО выбранного региона.
    Сезоны НЕ фильтруем — остаются как есть.
    Плюс: общая очистка хвостов под всеми таблицами.
    """
    region_key = str(region_name).strip().lower()
    los = sh_plan.api.ListObjects

    for tbl_name, meta in snapshot.items():
        headers = meta["headers"]
        df_all  = meta["df"].copy()
        anchor  = meta["anchor_rc"]

        reg_col = _find_region_col(headers)
        if reg_col is None or reg_col not in df_all.columns:
            # Таблица без столбца региона — не трогаем
            continue

        df_all[reg_col] = df_all[reg_col].astype(str)
        df_out = df_all[df_all[reg_col].str.strip().str.lower() == region_key].copy()

        # Сохраняем порядок столбцов
        df_out = df_out.reindex(columns=headers, fill_value="")

        # Перезаписываем таблицу и очищаем хвост
        lo = los.Item(tbl_name)
        _resize_and_write_table(sh_plan, lo, headers, df_out, anchor)

    # Доп. общая очистка хвостов по всему листу
    clean_tails_under_all_tables(sh_plan)


# --------------------------- Обновление сводных + валидация ---------------------------

def refresh_pivots(wb):
    """Надёжное обновление всех сводных таблиц и кэшей в книге."""
    # PivotCaches
    try:
        pcaches = wb.api.PivotCaches()
        for i in range(1, pcaches.Count + 1):
            try:
                pcaches.Item(i).Refresh()
            except Exception:
                pass
    except Exception:
        pass

    # PivotTables на всех листах
    for sht in wb.sheets:
        try:
            pivots = sht.api.PivotTables()
            for i in range(1, pivots.Count + 1):
                pt = pivots.Item(i)
                try:
                    pt.PivotCache().Refresh()
                except Exception:
                    pass
                try:
                    pt.RefreshTable()
                except Exception:
                    pass
        except Exception:
            continue

    # Полная перекалькуляция
    try:
        wb.api.Application.CalculateFullRebuild()
    except Exception:
        try:
            wb.api.Application.CalculateFull()
        except Exception:
            pass


def validate_plan_sheet_by_region(sh_plan, region_name):
    """
    Диагностика: проверяем, что в каждой таблице с колонкой региона остались только строки нужного региона.
    Печатаем предупреждения, если найдено что-то ещё.
    """
    region_key = str(region_name).strip().lower()
    los = sh_plan.api.ListObjects
    problems = 0

    for i in range(1, los.Count + 1):
        lo = los.Item(i)
        headers = [lo.ListColumns(j).Name for j in range(1, lo.ListColumns.Count + 1)]
        reg_col = _find_region_col(headers)
        if reg_col is None:
            continue

        if lo.DataBodyRange is None:
            continue

        vals = sh_plan.range(lo.DataBodyRange.Address).value
        if vals and not isinstance(vals[0], (list, tuple)):
            vals = [vals]
        df = pd.DataFrame(vals, columns=headers)

        # Если таблица пустая (одна техническая пустая строка)
        if df.dropna(how="all").empty:
            continue

        ser = df[reg_col].astype(str).str.strip().str.lower()
        uniq = sorted(set(ser.dropna().tolist()))
        # допустимо: либо пусто, либо только region_key
        if uniq and not (len(uniq) == 1 and uniq[0] == region_key):
            problems += 1
            print(Fore.YELLOW + f"⚠ Таблица '{lo.Name}': обнаружены регионы {uniq}, ожидался только '{region_key}'." + Style.RESET_ALL)

    if problems == 0:
        print(Fore.GREEN + "Проверка листа «План контрактации»: всё ОК, остались только строки выбранного региона." + Style.RESET_ALL)


# --------------------------- Основной сценарий ---------------------------

def separate_div_reg():
    start_time = time.time()

    divisions_order = [
        "Дивизион ДАЛЬНИЙ ВОСТОК", "Дивизион УРАЛ", "Дивизион СИБИРЬ",
        "Дивизион ПОВОЛЖЬЕ", "Дивизион ЦЕНТР", "Дивизион ЮГ"
    ]
    def get_division_order(d): return divisions_order.index(d) if d in divisions_order else len(divisions_order)

    region_division_map = {
        "Азербайджан":"СНГ","Узбекистан":"СНГ","Алтайский край":"Дивизион СИБИРЬ","Амурская область":"Дивизион ДАЛЬНИЙ ВОСТОК",
        "Астраханская область":"Дивизион ЮГ","Белгородская область":"Дивизион ЦЕНТР","Белоруссия":"СНГ",
        "Брянская область":"Дивизион ЦЕНТР","Владимирская область":"Дивизион ЦЕНТР",
        "Волгоградская область":"Дивизион ЮГ","Воронежская область":"Дивизион ЦЕНТР","Грузия":"СНГ",
        "ДНР":"Дивизион ЦЕНТР","Запорожье/Херсон":"Дивизион ЮГ","Иркутская область":"Дивизион СИБИРЬ",
        "Казахстан":"СНГ","Калининградская область":"Дивизион ЦЕНТР","Кемеровская область":"Дивизион СИБИРЬ",
        "Кировская область":"Дивизион ПОВОЛЖЬЕ","Краснодарский край 1":"Дивизион ЮГ",
        "Краснодарский край 2":"Дивизион ЮГ","Красноярский край":"Дивизион СИБИРЬ",
        "Курганская область":"Дивизион УРАЛ","Курская область":"Дивизион ЦЕНТР","Липецкая область":"Дивизион ЦЕНТР",
        "ЛНР":"Дивизион ЦЕНТР","ЛНР/ДНР":"Дивизион ЦЕНТР","Московская область":"Дивизион ЦЕНТР",
        "Нижегородская область":"Дивизион ПОВОЛЖЬЕ","Новосибирская область":"Дивизион СИБИРЬ",
        "Омская область":"Дивизион СИБИРЬ","Оренбургская область":"Дивизион ПОВОЛЖЬЕ",
        "Орловская область":"Дивизион ЦЕНТР","Пензенская область":"Дивизион ПОВОЛЖЬЕ",
        "Приморский край":"Дивизион ДАЛЬНИЙ ВОСТОК","Республика Башкортостан":"Дивизион ПОВОЛЖЬЕ",
        "Республика Дагестан":"Дивизион ЮГ","Республика Калмыкия":"Дивизион ЮГ","Республика Крым":"Дивизион ЮГ",
        "Республика Мордовия":"Дивизион ПОВОЛЖЬЕ","Республика Татарстан":"Дивизион ПОВОЛЖЬЕ",
        "Республика Чувашия":"Дивизион ПОВОЛЖЬЕ","Ростовская область 1":"Дивизион ЮГ",
        "Ростовская область 2":"Дивизион ЮГ","Рязанская область":"Дивизион ЦЕНТР",
        "Самарская область":"Дивизион ПОВОЛЖЬЕ","Саратовская область":"Дивизион ПОВОЛЖЬЕ",
        "Свердловская область":"Дивизион УРАЛ","Ставропольский край":"Дивизион ЮГ",
        "Тамбовская область":"Дивизион ЦЕНТР","Томская область":"Дивизион СИБИРЬ",
        "Тульская область":"Дивизион ЦЕНТР","Тюменская область":"Дивизион УРАЛ",
        "Ульяновская область":"Дивизион ПОВОЛЖЬЕ","Челябинская область":"Дивизион УРАЛ",
        "Беларусь":"СНГ","Армения":"СНГ"
    }
    allowed_yug = {"Республика Крым", "Республика Дагестан"}

    # Ищем исходный файл по маске (UNC)
    pattern = r'\\192.168.1.211\Аналитический центр\Отчёты\Быстрый старт\Отчет по продажам *.xlsx'
    files = glob.glob(pattern)
    if not files:
        print("Не найдено файлов по шаблону:", pattern)
        return
    filename = max(files, key=os.path.getctime)
    print("Обрабатывается файл:", filename)

    # Читаем листы с данными
    data_goods  = pd.read_excel(filename, sheet_name='data_товар', header=0)
    data_orders = pd.read_excel(filename, sheet_name='data_заказ', header=0)

    # Запускаем скрытый Excel
    app = xw.App(visible=False)
    api = app.api
    api.DisplayAlerts = False
    api.AskToUpdateLinks = False
    api.AlertBeforeOverwriting = False

    # Открываем книгу и подготавливаем листы
    wb = xw.Book(filename)
    sh_goods  = wb.sheets['data_товар']
    sh_orders = wb.sheets['data_заказ']
    sh_plan   = wb.sheets['План контрактации']

    # Делаем один "слепок" всех таблиц листа «План контрактации»
    plan_snapshot = snapshot_plan_tables(sh_plan)

    # Список регионов и маппинг дивизионов (исключаем СНГ и неразрешённые регионы Юга)
    div_regions_map = {}
    for r in data_orders['Наименование'].unique():
        div = region_division_map.get(r)
        if div and div != "СНГ":
            if div == "Дивизион ЮГ" and r not in allowed_yug:
                continue
            div_regions_map.setdefault(div, set()).add(r)

    processed_regions_by_div = {div: set() for div in div_regions_map.keys()}

    regions = [r for r in data_orders['Наименование'].unique()
               if region_division_map.get(r) and region_division_map[r] != "СНГ"]
    regions = [r for r in regions if not (region_division_map[r] == "Дивизион ЮГ" and r not in allowed_yug)]
    regions.sort(key=lambda r: get_division_order(region_division_map[r]))

    for region in regions:
        division = region_division_map[region]
        print(f"\n--- Обработка региона: {region} ({division}) ---")

        target_folder = os.path.join(r'\\192.168.1.211\торговый дом', division, region, 'Маркетинг')
        if os.path.exists(target_folder) and not os.path.isdir(target_folder):
            print("Путь есть как файл, пропускаем:", target_folder)
            continue
        os.makedirs(target_folder, exist_ok=True)

        # 1) Формируем и записываем таблицы data_товар и data_заказ по региону
        dg = data_goods[data_goods['Регион.Наименование'] == region].reset_index(drop=True)
        do = data_orders[data_orders['Наименование'] == region].reset_index(drop=True)

        # Удаляем существующие ListObjects на этих листах
        for i in range(sh_goods.api.ListObjects.Count, 0, -1):
            sh_goods.api.ListObjects(i).Delete()
        for i in range(sh_orders.api.ListObjects.Count, 0, -1):
            sh_orders.api.ListObjects(i).Delete()

        # Чистим только рабочие листы данных (НЕ лист планов!)
        safe_clear_sheet(sh_goods)
        safe_clear_sheet(sh_orders)

        # Пишем значения и снова создаём таблицы
        sh_goods.range('A1').options(index=False).value = dg
        sh_goods.tables.add(source=sh_goods.range('A1').expand(), name='data_товар')

        sh_orders.range('A1').options(index=False).value = do
        sh_orders.tables.add(source=sh_orders.range('A1').expand(), name='data_заказ')

        # 2) Фильтруем все таблицы на «План контрактации» ТОЛЬКО по региону (сезоны не трогаем)
        filter_plan_sheet_by_region_from_snapshot(sh_plan, plan_snapshot, region_name=region)

        # 3) Обновляем сводные (и кэши)
        refresh_pivots(wb)

        # 4) Автопроверка корректности фильтрации планов
        validate_plan_sheet_by_region(sh_plan, region_name=region)

        # 5) Сохраняем и копируем
        today = datetime.today().strftime('%m-%d')
        fname = f"Отчет по продажам {today} {region.replace('/', ' ')}.xlsx"
        full = os.path.join(os.getcwd(), fname)
        if os.path.exists(full):
            print("Файл уже есть, пропускаем:", full)
        else:
            wb.save(full)
            shutil.copy2(full, target_folder)
            print(Fore.GREEN + f"{region} – готово!" + Style.RESET_ALL)

        processed_regions_by_div[division].add(region)

    # Закрываем и чистим
    wb.close()
    app.kill()
    remove_reports_from_subfolders(os.getcwd())

    # Отправка писем по дивизионам, если обработаны все их регионы
    for division in div_regions_map.keys():
        if processed_regions_by_div[division] == div_regions_map[division]:
            threading.Thread(target=send_division_email, args=(division,), daemon=True).start()

    print("Общее время (мин):", (time.time() - start_time) / 60)


if __name__ == '__main__':
    separate_div_reg()
