import os
import pandas as pd
from datetime import datetime
import re


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


def find_date_and_income(folder_path):
    """
    Ищет в Excel файлах первую дату и суммы после ячейки со словом 'приход'
    """
    print(f"Поиск в папке: {folder_path}")
    print("=" * 80)

    # Находим все Excel файлы рекурсивно
    excel_files = find_excel_files(folder_path)

    if not excel_files:
        print("Файлы Excel не найдены в указанной папке и подпапках!")
        return

    # Собираем все результаты для табличного вывода
    all_results = []

    for file_path in excel_files:
        filename = os.path.basename(file_path)
        folder_name = os.path.dirname(file_path)

        try:
            # Читаем все листы Excel файла
            excel_file = pd.ExcelFile(file_path)

            file_date = None
            income_amounts = []

            for sheet_name in excel_file.sheet_names:
                # Читаем лист как есть (без преобразования типов для поиска дат)
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

                # Ищем первую дату в файле
                if file_date is None:
                    for col in df.columns:
                        for cell in df[col].dropna():
                            # Пробуем преобразовать в дату
                            if isinstance(cell, (datetime, pd.Timestamp)):
                                file_date = cell
                                break
                            elif isinstance(cell, str):
                                # Пробуем распарсить строку как дату
                                try:
                                    date_obj = pd.to_datetime(cell, errors='coerce')
                                    if not pd.isna(date_obj):
                                        file_date = date_obj
                                        break
                                except:
                                    pass
                        if file_date is not None:
                            break

                # Ищем слово 'приход' и числа после него
                for row_idx in range(len(df)):
                    for col_idx in range(len(df.columns)):
                        cell_value = df.iat[row_idx, col_idx]

                        # Проверяем, содержит ли ячейка слово 'приход' (регистронезависимо)
                        if (cell_value and
                                isinstance(cell_value, str) and
                                'приход' in str(cell_value).lower()):

                            # Ищем числа в соседних ячейках справа и снизу
                            # Справа от ячейки
                            if col_idx + 1 < len(df.columns):
                                right_cell = df.iat[row_idx, col_idx + 1]
                                amount = extract_number(right_cell)
                                if amount is not None:
                                    income_amounts.append(amount)

                            # Снизу от ячейки
                            if row_idx + 1 < len(df):
                                below_cell = df.iat[row_idx + 1, col_idx]
                                amount = extract_number(below_cell)
                                if amount is not None:
                                    income_amounts.append(amount)

                            # Диагональ справа-снизу
                            if (row_idx + 1 < len(df) and
                                    col_idx + 1 < len(df.columns)):
                                diag_cell = df.iat[row_idx + 1, col_idx + 1]
                                amount = extract_number(diag_cell)
                                if amount is not None:
                                    income_amounts.append(amount)

            # Форматируем дату
            if file_date:
                if isinstance(file_date, (datetime, pd.Timestamp)):
                    formatted_date = file_date.strftime('%Y-%m-%d')
                else:
                    formatted_date = str(file_date)
            else:
                formatted_date = "Не найдена"

            # Убираем дубликаты сумм
            income_amounts = list(set(income_amounts))

            # Добавляем результаты
            if income_amounts:
                for amount in income_amounts:
                    all_results.append({
                        'Файл': filename,
                        'Папка': folder_name,
                        'Дата': formatted_date,
                        'Сумма прихода': f"{amount:,.2f}".replace(',', ' ').replace('.', ',')
                    })
            else:
                all_results.append({
                    'Файл': filename,
                    'Папка': folder_name,
                    'Дата': formatted_date,
                    'Сумма прихода': "Не найдена"
                })

        except Exception as e:
            all_results.append({
                'Файл': filename,
                'Папка': folder_name,
                'Дата': "Ошибка",
                'Сумма прихода': f"Ошибка: {str(e)}"
            })

    # Выводим таблицу
    if all_results:
        # Создаем DataFrame для красивого вывода
        df_results = pd.DataFrame(all_results)

        # Настраиваем отображение pandas для красивого вывода
        pd.set_option('display.max_rows', None)
        pd.set_option('display.width', None)
        pd.set_option('display.max_colwidth', 50)

        print(f"\nНайдено файлов: {len(excel_files)}")
        print("\nРЕЗУЛЬТАТЫ ПОИСКА:")
        print("=" * 100)
        print(df_results.to_string(index=False))
        print("=" * 100)

        # Статистика
        total_files = len(set(result['Файл'] for result in all_results))
        total_amounts = len([result for result in all_results if
                             result['Сумма прихода'] != "Не найдена" and "Ошибка" not in result['Сумма прихода']])

        print(f"\nСТАТИСТИКА:")
        print(f"Обработано файлов: {total_files}")
        print(f"Найдено сумм прихода: {total_amounts}")


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