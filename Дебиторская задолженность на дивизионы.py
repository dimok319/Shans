import pandas as pd
import xlwings as xw
import os
import glob
import time
import shutil  # для переноса старых файлов в архив
from datetime import datetime
from colorama import init, Fore  # для цветного вывода в консоль

# Инициализация colorama
init(autoreset=True)


def payment_by_division():
    start_time = time.time()

    # --- Шаг 1. Находим последний сохранённый файл в папке ---
    folder_path = r"\\192.168.1.211\Аналитический центр\Отчёты\ДЗ\Контроль сроков оплаты"
    list_of_files = glob.glob(os.path.join(folder_path, 'Отчет по ДЗ*.xlsx'))
    source_file = max(list_of_files, key=os.path.getmtime)
    print(f"{Fore.GREEN}Выбранный файл для обработки: {source_file}")

    # --- Шаг 2. Словарь соответствия "Регион -> Дивизион" ---
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

    # --- Шаг 3. Читаем исходные данные в DataFrame ---
    df = pd.read_excel(source_file, sheet_name='TDSheet', header=0)
    # Объединяем данные для Алтайского края 1 и Алтайского края 2
    df['Регион.Наименование'] = df['Регион.Наименование'].replace(
        {'Алтайский край 1': 'Алтайский край', 'Алтайский край 2': 'Алтайский край'})
    # Переопределяем "Дивизион" по словарю
    df['Дивизион'] = df['Регион.Наименование'].map(region_division_map)
    df = df.dropna(subset=['Дивизион'])

    # --- Шаг 4. Используем полный список дивизионов из словаря (кроме СНГ) ---
    all_divisions = list(set(region_division_map.values()))
    if "СНГ" in all_divisions:
        all_divisions.remove("СНГ")

    root_dir = r"\\192.168.1.211\торговый дом"
    new_date_str = datetime.today().strftime('%Y%m%d')

    # Для каждого дивизиона открываем новый экземпляр книги, чтобы не «загрязнять» данные предыдущей итерацией
    for division in all_divisions:
        division_start_time = time.time()

        # Открываем исходную книгу заново для текущей итерации
        app_excel = xw.App(visible=False)
        wb = xw.Book(source_file)
        sh_data = wb.sheets['TDSheet']

        # Формируем путь для сохранения файла по дивизиону:
        # \\192.168.1.211\торговый дом\<division>\Дебиторская задолженность\Отчеты
        division_dir = os.path.join(root_dir, division)
        debt_dir = os.path.join(division_dir, "Дебиторская задолженность")
        reports_dir = os.path.join(debt_dir, "Отчеты")
        archive_dir = os.path.join(reports_dir, "Архив")
        os.makedirs(archive_dir, exist_ok=True)

        # Переносим старые отчёты в архив (работаем с файлами в папке reports_dir)
        old_reports = glob.glob(os.path.join(reports_dir, '*_Отчёт по ДЗ_*.xlsx'))
        for rep in old_reports:
            if '~$' in rep:
                continue
            archive_file = os.path.join(archive_dir, os.path.basename(rep))
            filename = os.path.basename(rep)
            file_date_str = filename[:8]
            if os.path.exists(archive_file) and file_date_str != new_date_str:
                try:
                    os.remove(rep)
                    print(f"{Fore.YELLOW}{rep} удалён из основной папки, так как в архиве уже есть файл с этой датой.")
                except Exception as e:
                    print(f"{Fore.RED}Не удалось удалить {rep}: {e}")
            else:
                try:
                    shutil.move(rep, archive_dir)
                    print(f"{Fore.GREEN}{rep} перемещён в архив.")
                except Exception as e:
                    print(f"{Fore.RED}Не удалось переместить {rep} в архив: {e}")

        # Фильтруем данные по текущему дивизиону
        df_div = df.loc[df['Дивизион'] == division].reset_index(drop=True)

        # Перед записью новых данных удаляем существующую таблицу "TDSheet", если она есть
        try:
            sh_data.api.ListObjects("TDSheet").Delete()
        except Exception:
            pass

        sh_data.clear_contents()
        sh_data.range('A1').options(index=False).value = df_div

        # Если есть данные, создаём таблицу
        if not df_div.empty:
            sh_data.tables.add(source=sh_data.range('A1').expand(), name='TDSheet')

        wb.api.RefreshAll()
        try:
            wb.sheets['Соблюдение оплаты'].activate()
        except Exception:
            pass

        # Формируем имя файла: <дата>_Отчёт по ДЗ_<название дивизиона>
        division_name_for_file = division.replace("Дивизион ", "")
        report_name = f"{new_date_str}_Отчёт по ДЗ_{division_name_for_file}.xlsx"
        report_path = os.path.join(reports_dir, report_name)

        if not os.path.exists(report_path):
            try:
                wb.save(report_path)
            except Exception as e:
                print(f"{Fore.RED}Ошибка при сохранении файла для дивизиона {division}: {e}")
        else:
            print(f"{Fore.YELLOW}Файл для дивизиона {division} уже существует, пропускаем сохранение.")

        division_end_time = time.time()
        print(f"{Fore.GREEN}{division} - обработан. Время обработки: {(division_end_time - division_start_time):.2f} секунд")
        wb.close()
        app_excel.kill()

    end_time = time.time()
    print(f'Общее время работы программы: {(end_time - start_time) / 60:.2f} минут(ы).')


if __name__ == "__main__":
    payment_by_division()
