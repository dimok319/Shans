import os
import pandas as pd
from datetime import datetime
import re
import warnings

warnings.filterwarnings('ignore')


def find_excel_files(folder_path):
    """
    Рекурсивно ищет все Excel файлы в папке и подпапках
    """
    excel_files = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(('.xlsx', '.xls')):
                full_path = os.path.join(root, file)
                excel_files.append(full_path)

    return excel_files


def extract_date_from_filename(filename):
    """
    Извлекает дату из названия файла
    """
    # Паттерны для поиска дат в названиях файлов
    date_patterns = [
        r'(\d{1,2}\s+\w+\s+\d{4})',  # "30 января 2025"
        r'(\d{1,2}\.\d{1,2}\.\d{4})',  # "30.01.2025"
        r'(\d{1,2}-\d{1,2}-\d{4})',  # "30-01-2025"
        r'(\d{4}-\d{1,2}-\d{1,2})',  # "2025-01-30"
    ]

    for pattern in date_patterns:
        match = re.search(pattern, filename)
        if match:
            date_str = match.group(1)
            try:
                # Пробуем разные форматы дат
                if re.match(r'\d{1,2}\s+\w+\s+\d{4}', date_str):
                    # Русская дата "30 января 2025"
                    months_ru = {
                        'января': 1, 'февраля': 2, 'марта': 3, 'апреля': 4,
                        'мая': 5, 'июня': 6, 'июля': 7, 'августа': 8,
                        'сентября': 9, 'октября': 10, 'ноября': 11, 'декабря': 12
                    }
                    day, month_ru, year = date_str.split()
                    month = months_ru.get(month_ru.lower())
                    if month:
                        return datetime(int(year), month, int(day))
                else:
                    # Стандартные форматы дат
                    date_obj = pd.to_datetime(date_str, dayfirst=True, errors='coerce')
                    if not pd.isna(date_obj):
                        return date_obj
            except:
                continue
    return None


def find_date_and_income(folder_path):
    """
    Ищет в Excel файлах первую дату и суммы в правой ячейке от ячейки со словом 'приход'
    Оставляет максимальную сумму для каждой даты
    """
    print(f"Поиск в папке: {folder_path}")
    print("=" * 80)

    # Находим все Excel файлы рекурсивно
    excel_files = find_excel_files(folder_path)

    if not excel_files:
        print("Файлы Excel не найдены в указанной папке и подпапках!")
        return

    # Собираем все результаты для последующей обработки
    all_results = []

    for file_path in excel_files:
        filename = os.path.basename(file_path)
        folder_name = os.path.dirname(file_path)

        try:
            # Читаем все листы Excel файла
            excel_file = pd.ExcelFile(file_path)

            file_date = None
            income_amounts = []

            # Сначала пробуем извлечь дату из названия файла
            file_date = extract_date_from_filename(filename)
            date_source = "из названия файла"

            # Если не нашли в названии, ищем в содержимом файла
            if file_date is None:
                for sheet_name in excel_file.sheet_names:
                    # Читаем лист как есть (без преобразования типов для поиска дат)
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

                    # Ищем первую дату в файле
                    for col in df.columns:
                        for cell in df[col].dropna():
                            # Пробуем преобразовать в дату
                            if isinstance(cell, (datetime, pd.Timestamp)):
                                file_date = cell
                                date_source = "из содержимого файла"
                                break
                            elif isinstance(cell, str):
                                # Пробуем распарсить строку как дату
                                try:
                                    date_obj = pd.to_datetime(cell, dayfirst=True, errors='coerce')
                                    if not pd.isna(date_obj):
                                        file_date = date_obj
                                        date_source = "из содержимого файла"
                                        break
                                except:
                                    pass
                        if file_date is not None:
                            break
                    if file_date is not None:
                        break

            # Ищем слово 'приход' и числа в правой ячейке
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

                for row_idx in range(len(df)):
                    for col_idx in range(len(df.columns)):
                        cell_value = df.iat[row_idx, col_idx]

                        # Проверяем, содержит ли ячейка слово 'приход' (регистронезависимо)
                        if (cell_value and
                                isinstance(cell_value, str) and
                                'приход' in str(cell_value).lower()):

                            # Ищем число ТОЛЬКО в соседней правой ячейке
                            if col_idx + 1 < len(df.columns):
                                right_cell = df.iat[row_idx, col_idx + 1]

                                # Проверяем, что ячейка не пустая
                                if right_cell is not None and not pd.isna(right_cell):
                                    amount = extract_number(right_cell)
                                    if amount is not None:
                                        income_amounts.append(amount)

            # Форматируем дату для группировки
            if file_date:
                if isinstance(file_date, (datetime, pd.Timestamp)):
                    date_key = file_date.strftime('%Y-%m-%d')  # Ключ для группировки
                    formatted_date = file_date.strftime('%Y-%m-%d')
                    date_display = f"{formatted_date} ({date_source})"
                else:
                    date_key = str(file_date)
                    formatted_date = str(file_date)
                    date_display = f"{formatted_date} ({date_source})"
            else:
                date_key = "Не найдена"
                date_display = "Не найдена"

            # Находим максимальную сумму прихода для этого файла
            max_amount = max(income_amounts) if income_amounts else None

            # Добавляем результат
            if max_amount is not None:
                all_results.append({
                    'Файл': filename,
                    'Папка': folder_name,
                    'Дата': date_display,
                    'Дата_ключ': date_key,
                    'Сумма прихода': max_amount,
                    'Сумма прихода (формат)': f"{max_amount:,.2f}".replace(',', ' ').replace('.', ',')
                })
            else:
                all_results.append({
                    'Файл': filename,
                    'Папка': folder_name,
                    'Дата': date_display,
                    'Дата_ключ': date_key,
                    'Сумма прихода': None,
                    'Сумма прихода (формат)': "Не найдена"
                })

        except Exception as e:
            all_results.append({
                'Файл': filename,
                'Папка': folder_name,
                'Дата': "Ошибка",
                'Дата_ключ': "Ошибка",
                'Сумма прихода': None,
                'Сумма прихода (формат)': f"Ошибка: {str(e)}"
            })

    # Группируем по дате и оставляем максимальную сумму для каждой даты
    grouped_results = {}
    for result in all_results:
        date_key = result['Дата_ключ']
        amount = result['Сумма прихода']

        if date_key not in grouped_results:
            grouped_results[date_key] = result
        else:
            # Если для этой даты уже есть запись, сравниваем суммы
            current_amount = grouped_results[date_key]['Сумма прихода']
            if amount is not None and (current_amount is None or amount > current_amount):
                grouped_results[date_key] = result

    # Преобразуем обратно в список для вывода
    final_results = list(grouped_results.values())

    # Выводим таблицу
    if final_results:
        # Создаем DataFrame для красивого вывода
        df_results = pd.DataFrame([{
            'Файл': r['Файл'],
            'Папка': r['Папка'],
            'Дата': r['Дата'],
            'Сумма прихода': r['Сумма прихода (формат)']
        } for r in final_results])

        # Сортируем по дате
        try:
            df_results = df_results.sort_values('Дата')
        except:
            pass  # Если сортировка по дате не удалась, оставляем как есть

        # Настраиваем отображение pandas для красивого вывода
        pd.set_option('display.max_rows', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 50)

        print(f"\nНайдено файлов: {len(excel_files)}")
        print("\nРЕЗУЛЬТАТЫ ПОИСКА (максимальная сумма для каждой даты):")
        print("=" * 120)
        print(df_results.to_string(index=False))
        print("=" * 120)

        # Статистика
        total_files = len(excel_files)
        files_with_dates = len(
            [r for r in final_results if "Не найдена" not in r['Дата'] and "Ошибка" not in r['Дата']])
        files_with_amounts = len([r for r in final_results if
                                  r['Сумма прихода'] is not None and "Ошибка" not in r['Сумма прихода (формат)']])
        unique_dates = len([r for r in final_results if "Не найдена" not in r['Дата'] and "Ошибка" not in r['Дата']])

        print(f"\nСТАТИСТИКА:")
        print(f"Обработано файлов: {total_files}")
        print(f"Уникальных дат: {unique_dates}")
        print(f"Файлов с найденными датами: {files_with_dates}")
        print(f"Файлов с найденными суммами: {files_with_amounts}")


def extract_number(value):
    """
    Извлекает число из значения ячейки
    """
    if value is None or pd.isna(value):
        return None

    # Если уже число
    if isinstance(value, (int, float)):
        return float(value)

    # Если строка, пытаемся извлечь число
    if isinstance(value, str):
        # Убираем пробелы и запятые (для разделителей тысяч)
        cleaned = value.replace(' ', '').replace(',', '.')

        # Ищем числа с плавающей точкой
        match = re.search(r'(\d+[.,]?\d*)', cleaned)
        if match:
            try:
                return float(match.group(1).replace(',', '.'))
            except ValueError:
                pass

    return None


# Использование
if __name__ == "__main__":
    folder_path = r"C:\Users\dpyshnograev\Desktop\апвапвап"

    if not os.path.exists(folder_path):
        print(f"Папка не существует: {folder_path}")
    else:
        find_date_and_income(folder_path)