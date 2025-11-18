import pdfplumber
import pandas as pd
import os

# Путь к файлу PDF
pdf_path = r'C:\Users\nkazakov\Downloads\Цены Щелково.pdf'

# Извлекаем имя файла без расширения
pdf_filename = os.path.splitext(os.path.basename(pdf_path))[0]
output_path = rf'C:\\Users\\nkazakov\\Downloads\\{pdf_filename}.xlsx'

# Определяем заголовки, которые должны быть в файле Excel
expected_headers = [
    "Препарат", "Действующие вещества", "Норма расхода, кг(л)/га(т)",
    "Тарная упаковка", "Цена с НДС"
]

# Создаем список для хранения всех данных
all_data = []
previous_entry = None  # Хранение предыдущей строки

# Открываем PDF-файл
with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        tables = page.extract_tables()
        for table in tables:
            if table:
                for row in table:
                    row = [cell.strip() if cell else None for cell in row]

                    # Если первая ячейка пустая, продолжаем предыдущую строку
                    if row[0] is None and previous_entry:
                        new_entry = previous_entry[:]
                        for i in range(1, len(row)):
                            if row[i]:
                                new_entry[i] = row[i]  # Заполняем данными текущей строки
                        all_data.append(new_entry)  # Добавляем как новую строку
                    else:
                        if previous_entry:
                            all_data.append(previous_entry[:])  # Добавляем завершенную строку
                        previous_entry = row[:len(expected_headers)] + [None] * (len(expected_headers) - len(row))

                if previous_entry:
                    all_data.append(previous_entry[:])  # Добавляем последнюю строку

# Создаем DataFrame
df = pd.DataFrame(all_data, columns=expected_headers)

# Сохраняем в Excel
df.to_excel(output_path, sheet_name='Data', index=False)

print(f"Таблицы успешно извлечены и сохранены в файл {output_path}")
