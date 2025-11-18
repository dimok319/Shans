import pandas as pd
import xlwings as xw
import os
import shutil
import glob
import time
from datetime import datetime

import colorama
from colorama import Fore, Style

colorama.init(autoreset=True)


def remove_reports_from_subfolders(start_dir):
    """Удаляет файлы, начинающиеся на 'Отчет по продажам' за текущую дату."""
    date_prefix = datetime.today().strftime('%m-%d')
    for dirpath, _, filenames in os.walk(start_dir):
        for filename in [f for f in filenames if
                         f.startswith("Отчет по продажам" + date_prefix) and f.endswith(".xlsx")]:
            try:
                os.remove(os.path.join(dirpath, filename))
            except Exception as e:
                print(f"Ошибка при удалении файла {filename}: {e}")


def safe_clear_sheet(sheet, attempts=5, delay=0.5):
    """
    Попытка очистить используемый диапазон листа с повторными попытками.
    """
    for i in range(attempts):
        try:
            sheet.api.UsedRange.ClearContents()
            return
        except Exception as e:
            time.sleep(delay)
    print(f"Не удалось очистить лист {sheet.name} после {attempts} попыток.")


def update_or_create_plan_table(sh_plan, table_name, filtered_df):
    """
    Обновляет или создаёт таблицу (ListObject) с именем table_name на листе sh_plan.
    Заголовок таблицы находится в ячейке A2, а данные начинаются ниже.
    Если таблица существует – её размер изменяется, иначе создаётся новая.
    Параметры передаются позиционно, так как COM‑метод не принимает именованные аргументы.
    """
    # Полностью очищаем лист с планами (без RefreshAll, чтобы не сбрасывать наши данные)
    safe_clear_sheet(sh_plan)

    start_cell = sh_plan.range("A2")
    cols = list(filtered_df.columns)
    ncols = len(cols)
    nrows = filtered_df.shape[0]
    new_height = nrows + 1  # 1 строка заголовка + данные
    new_range = start_cell.resize(new_height, ncols)
    new_data = [cols] + filtered_df.values.tolist()
    new_range.value = new_data

    try:
        # Пытаемся получить существующую таблицу
        lo = sh_plan.api.ListObjects(table_name)
        lo.Resize(new_range.api)
    except Exception as e:
        try:
            # Если таблицы нет, создаём новую с позиционными аргументами:
            # (SourceType=1, SourceData=new_range.api, LinkSource=False, XlListObjectHasHeaders=1, Destination=new_range.api)
            lo = sh_plan.api.ListObjects.Add(1, new_range.api, False, 1, new_range.api)
            lo.Name = table_name
        except Exception as ex:
            print(f"Ошибка при создании таблицы '{table_name}': {ex}")


def separate_div_reg():
    start_time = time.time()

    # Порядок дивизионов для сортировки регионов
    divisions_order = [
        "Дивизион ДАЛЬНИЙ ВОСТОК",
        "Дивизион УРАЛ",
        "Дивизион СИБИРЬ",
        "Дивизион ПОВОЛЖЬЕ",
        "Дивизион ЦЕНТР",
        "Дивизион ЮГ"
    ]

    def get_division_order(div_name):
        return divisions_order.index(div_name) if div_name in divisions_order else len(divisions_order)

    # Сопоставление регионов с дивизионами (СНГ не обрабатываем)
    region_division_map = {
        "Азербайджан": "СНГ",
        "Алтайский край": "Дивизион СИБИРЬ",
        "Амурская область": "Дивизион ДАЛЬНИЙ ВОСТОК",
        "Астраханская область": "Дивизион ЮГ",
        "Белгородская область": "Дивизион ЦЕНТР",
        "Белоруссия": "СНГ",
        "Брянская область": "Дивизион ЦЕНТР",
        "Владимирская область": "Дивизион ЦЕНТР",
        "Волгоградская область": "Дивизион ЮГ",
        "Воронежская область": "Дивизион ЦЕНТР",
        "Грузия": "СНГ",
        "ДНР": "Дивизион ЦЕНТР",
        "Запорожье/Херсон": "Дивизион ЮГ",
        "Иркутская область": "Дивизион СИБИРЬ",
        "Казахстан": "СНГ",
        "Калининградская область": "Дивизион ЦЕНТР",
        "Кемеровская область": "Дивизион СИБИРЬ",
        "Кировская область": "Дивизион ПОВОЛЖЬЕ",
        "Краснодарский край 1": "Дивизион ЮГ",
        "Краснодарский край 2": "Дивизион ЮГ",
        "Красноярский край": "Дивизион СИБИРЬ",
        "Курганская область": "Дивизион УРАЛ",
        "Курская область": "Дивизион ЦЕНТР",
        "Липецкая область": "Дивизион ЦЕНТР",
        "ЛНР": "Дивизион ЦЕНТР",
        "ЛНР/ДНР": "Дивизион ЦЕНТР",
        "Московская область": "Дивизион ЦЕНТР",
        "Нижегородская область": "Дивизион ПОВОЛЖЬЕ",
        "Новосибирская область": "Дивизион СИБИРЬ",
        "Омская область": "Дивизион СИБИРЬ",
        "Оренбургская область": "Дивизион ПОВОЛЖЬЕ",
        "Орловская область": "Дивизион ЦЕНТР",
        "Пензенская область": "Дивизион ПОВОЛЖЬЕ",
        "Приморский край": "Дивизион ДАЛЬНИЙ ВОСТОК",
        "Республика Башкортостан": "Дивизион ПОВОЛЖЬЕ",
        "Республика Дагестан": "Дивизион ЮГ",
        "Республика Калмыкия": "Дивизион ЮГ",
        "Республика Крым": "Дивизион ЮГ",
        "Республика Мордовия": "Дивизион ПОВОЛЖЬЕ",
        "Республика Татарстан": "Дивизион ПОВОЛЖЬЕ",
        "Республика Чувашия": "Дивизион ПОВОЛЖЬЕ",
        "Ростовская область 1": "Дивизион ЮГ",
        "Ростовская область 2": "Дивизион ЮГ",
        "Рязанская область": "Дивизион ЦЕНТР",
        "Самарская область": "Дивизион ПОВОЛЖЬЕ",
        "Саратовская область": "Дивизион ПОВОЛЖЬЕ",
        "Свердловская область": "Дивизион УРАЛ",
        "Ставропольский край": "Дивизион ЮГ",
        "Тамбовская область": "Дивизион ЦЕНТР",
        "Томская область": "Дивизион СИБИРЬ",
        "Тульская область": "Дивизион ЦЕНТР",
        "Тюменская область": "Дивизион УРАЛ",
        "Ульяновская область": "Дивизион ПОВОЛЖЬЕ",
        "Челябинская область": "Дивизион УРАЛ",
        "Беларусь": "СНГ",
        "Армения": "СНГ"
    }
    # Для дивизиона ЮГ обрабатываем только "Республика Крым" и "Республика Дагестан"
    allowed_yug_regions = {"Республика Крым", "Республика Дагестан"}

    files_pattern = r'\\192.168.1.211\Аналитический центр\Отчёты\Быстрый старт\Отчет по продажам *.xlsx'
    list_of_files = glob.glob(files_pattern)
    if not list_of_files:
        print("Не найдено файлов по шаблону:", files_pattern)
        return
    filename = max(list_of_files, key=os.path.getctime)
    print(f"Обрабатывается файл: {filename}")

    # Считываем исходные данные
    data_goods = pd.read_excel(filename, sheet_name='data_товар', header=0)
    data_orders = pd.read_excel(filename, sheet_name='data_заказ', header=0)
    plan_50 = pd.read_excel(filename, sheet_name='Планы скорректированные в минус', usecols='A:R', header=1)
    plan_shipment = pd.read_excel(filename, sheet_name='Планы скорректированные в минус', usecols='A:R', header=1)

    if 'ОП' in plan_50.columns:
        region_col = 'ОП'
    elif 'Регион' in plan_50.columns:
        region_col = 'Регион'
    else:
        print("В плановых данных отсутствует столбец 'ОП' или 'Регион'.")
        return

    print("Уникальные значения столбца для региона в планах:", plan_50[region_col].unique())

    # Создаём экземпляр Excel в фоновом режиме
    app_excel = xw.App(visible=False)
    wb = xw.Book(filename)
    sh_data_goods = wb.sheets['data_товар']
    sh_data_orders = wb.sheets['data_заказ']
    try:
        sh_plan = wb.sheets['Планы скорректированные в минус']
    except Exception as e:
        print(f"Не удалось получить лист 'Планы скорректированные в минус': {e}")
        wb.close()
        app_excel.kill()
        return

    all_regions = data_orders['Наименование'].unique()
    all_regions = [r for r in all_regions if
                   region_division_map.get(r, None) is not None and region_division_map[r] != "СНГ"]
    filtered_regions = []
    for r in all_regions:
        if region_division_map[r] == "Дивизион ЮГ":
            if r in allowed_yug_regions:
                filtered_regions.append(r)
        else:
            filtered_regions.append(r)
    all_regions = filtered_regions
    all_regions.sort(key=lambda reg: get_division_order(region_division_map[reg]))

    for region in all_regions:
        division = region_division_map[region]
        region_start = time.time()
        print(f"\n--- Обработка региона: {region} (Дивизион: {division}) ---")

        base_path = r'\\192.168.1.211\торговый дом'
        target_folder = os.path.join(base_path, division, region, 'Маркетинг')
        # Если целевой путь существует как файл, пропускаем регион
        if os.path.exists(target_folder) and not os.path.isdir(target_folder):
            print(f"Путь {target_folder} существует как файл. Пропускаем регион {region}.")
            continue
        os.makedirs(target_folder, exist_ok=True)

        dg_region = data_goods.loc[data_goods['Регион.Наименование'] == region].reset_index(drop=True)
        do_region = data_orders.loc[data_orders['Наименование'] == region].reset_index(drop=True)

        filtered_plan = plan_50[
            (plan_50[region_col].astype(str).str.strip().str.lower() == region.strip().lower()) &
            (pd.to_numeric(plan_50['Сезон'], errors='coerce') == 25)
            ].reset_index(drop=True)
        print(f"Для региона '{region}' найдено строк в плане: {len(filtered_plan)}")

        safe_clear_sheet(sh_data_goods)
        safe_clear_sheet(sh_data_orders)
        sh_data_goods.range('A1').options(index=False).value = dg_region
        sh_data_goods.tables.add(source=sh_data_goods.range('A1').expand(), name='data_товар')
        sh_data_orders.range('A1').options(index=False).value = do_region
        sh_data_orders.tables.add(source=sh_data_orders.range('A1').expand(), name='data_заказ')

        # Обновляем или создаём таблицу "Планы_Update_АГХ" на листе "Планы скорректированные в минус"
        update_or_create_plan_table(sh_plan, "Планы_Update_АГХ", filtered_plan)

        wb.api.RefreshAll()

        today_str = datetime.today().strftime('%m-%d')
        new_filename = f"Отчет по продажам {today_str} {region.replace('/', ' ')}.xlsx"
        target_file = os.path.join(os.getcwd(), new_filename)
        if os.path.exists(target_file):
            print(f"Файл {target_file} уже существует, пропускаем регион {region}.")
            continue

        try:
            wb.save(target_file)
        except Exception as e:
            print(f"Ошибка сохранения для региона {region}: {e}")
            wb.close()
            app_excel.kill()
            return

        try:
            shutil.copy2(target_file, target_folder)
        except Exception as e:
            print(f"Ошибка копирования файла для региона {region}: {e}")

        region_time = time.time() - region_start
        print(f"Время обработки региона {region}: {region_time:.2f} сек.")
        print(Fore.GREEN + f"{region} - успешно обработан!" + Style.RESET_ALL)

    wb.close()
    app_excel.kill()
    remove_reports_from_subfolders(os.getcwd())
    total_time = time.time() - start_time
    print(f"\nОбщее время работы программы: {total_time / 60:.2f} мин.")


if __name__ == '__main__':
    separate_div_reg()
