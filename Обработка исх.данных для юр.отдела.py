import pandas as pd

# Загрузка данных из Excel
file_path = r"C:\Users\nkazakov\Downloads\тест для Питона.xlsx"  # Указанный путь к файлу
sheet_name = 0  # Можно указать нужный лист, если их несколько

df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)

# Столбцы, которые нужно разложить
columns_to_split = [
    "Отгрузки по датам", "Оплаты по датам", "Плановые погашения по датам", "УПД", "Данные спецификаций"
]

# Столбцы, которые должны остаться неизменными
columns_to_keep = [
    "Контрагент", "Договор", "Дата договора", "Номер", "ИНН", "Процент коммерческого кредита",
    "Юр адрес", "Остаток задолженности", "Сумма отгружено", "Сумма оплачено", "Сумма заказано"
]

# Разбиваем каждую ячейку на список значений
for col in columns_to_split:
    df[col] = df[col].fillna("").apply(lambda x: x.split("\n") if isinstance(x, str) else [])

# Определяем максимальное количество строк для каждого контрагента
grouped = df.groupby(columns_to_keep, dropna=False, sort=False)
expanded_data = []

for base_values, group in grouped:
    max_rows = max(
        [group[col].apply(len).max() for col in columns_to_split if col in group and not group[col].isna().all()],
        default=0)
    base_dict = dict(zip(columns_to_keep, base_values))
    split_values = {col: sum(group[col].tolist(), []) if col in group else [] for col in columns_to_split}

    for i in range(max_rows):
        new_row = base_dict.copy()
        for col in columns_to_split:
            if i < len(split_values[col]):
                value = split_values[col][i]
                if col == "УПД":
                    parts = value.split(";")
                    new_row[f"{col} - дата"] = parts[0].strip() if len(parts) > 0 else ""
                    new_row[f"{col} - номер"] = parts[1].strip() if len(parts) > 1 else ""
                    new_row[f"{col} - сумма"] = parts[2].strip() if len(parts) > 2 else ""
                else:
                    date, amount = value.split(";") if ";" in value else (value, "")
                    new_row[f"{col} - дата"] = date.strip()
                    new_row[f"{col} - сумма"] = amount.strip()
            else:
                new_row[f"{col} - дата"] = ""
                new_row[f"{col} - сумма"] = ""
                if col == "УПД":
                    new_row[f"{col} - номер"] = ""
        expanded_data.append(new_row)

# Создание итогового DataFrame
expanded_df = pd.DataFrame(expanded_data)

# Проверяем, есть ли нужные столбцы перед удалением пустых строк
existing_columns = [f"{col} - дата" for col in columns_to_split if f"{col} - дата" in expanded_df.columns]
if existing_columns:
    expanded_df = expanded_df.dropna(how='all', subset=existing_columns)

# Проверяем, не пустой ли датафрейм перед сохранением
if not expanded_df.empty:
    output_file = r"C:\Users\nkazakov\Downloads\Обработанный исходник.xlsx"
    expanded_df.to_excel(output_file, index=False)
    print(f"Обработанные данные сохранены в {output_file}")
else:
    print("Ошибка: после обработки получен пустой файл. Проверьте исходные данные.")
