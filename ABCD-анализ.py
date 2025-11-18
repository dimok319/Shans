import pandas as pd
import numpy as np
import time
import warnings

# Подавляем предупреждения о deprecated поведении groupby.apply
warnings.filterwarnings("ignore", category=DeprecationWarning)

start_time = time.time()

# Путь к исходному файлу
path_in = r'\\192.168.1.211\Аналитический центр\Отчёты\Управление клиентской базой\Управление клиентской базой ABC Исх.xlsx'
path_out = r'\\192.168.1.211\Аналитический центр\Отчёты\Управление клиентской базой\Управление клиентской базой ABC_результат_new03_10.xlsx'

# 1. Обработка данных с листа "data_товар"
sheet_data = 'data_товар'
df_data = pd.read_excel(path_in, sheet_name=sheet_data, engine='openpyxl')

# Приводим типы столбцов
df_data['Дивизион'] = df_data['Дивизион'].astype(str)
df_data['Регион.Наименование'] = df_data['Регион.Наименование'].astype(str)
df_data['Торговый представитель'] = df_data['Торговый представитель'].astype(str)
df_data['Клиент'] = df_data['Клиент'].astype(str)

# Формируем столбец "Канал_продаж"
df_data['Канал_продаж'] = np.where(df_data['Дивизион'] == "СНГ", "СНГ", df_data['Канал продаж'])
df_data.drop(columns=['Канал продаж'], inplace=True)
df_data.rename(columns={'Канал_продаж': 'Канал продаж'}, inplace=True)

# Извлекаем название месяца
df_data['Месяз заказ'] = pd.to_datetime(df_data['Дата']).dt.month_name()

# Расчёт "% отклонения от МРЦ"
df_data['% отклонения от МРЦ'] = df_data['Цена'] / df_data['Цена по прайсу'] - 1

# Замена значений в "Регион.Наименование"
df_data['Регион.Наименование'] = df_data['Регион.Наименование'].replace({
    'Алтайский край 1': 'Алтайский край',
    'Алтайский край 2': 'Алтайский край'
})

# Определяем новый "Дивизион" по региону
def get_division(row):
    region = row["Регион.Наименование"]
    if region in ["Амурская область", "Приморский край"]:
        return "Дивизион ДАЛЬНИЙ ВОСТОК"
    elif region in ["Челябинская область", "Тюменская область", "Курганская область", "Свердловская область"]:
        return "Дивизион УРАЛ"
    elif region in ["Республика Башкортостан", "Республика Татарстан"]:
        return "Дивизион ПОВОЛЖЬЕ"
    elif region in ["Алтайский край", "Кемеровская область", "Красноярский край",
                    "Новосибирская область", "Омская область", "Томская область"]:
        return "Дивизион СИБИРЬ"
    elif region == "Тамбовская область":
        return "Дивизион ЦЕНТР"
    else:
        return row["Дивизион"]

df_data['Дивизион New'] = df_data.apply(get_division, axis=1)
df_data.drop(columns=['Дивизион'], inplace=True)
df_data.rename(columns={'Дивизион New': 'Дивизион', 'Регион.Наименование': 'Регион'}, inplace=True)

df_data['Дата'] = pd.to_datetime(df_data['Дата']).dt.date

# Заполнение отсутствующих значений в "Цена по прайсу"
df_data["Цена по прайсу"] = df_data.groupby(["Заказ клиента", "Номенклатура.Позиция классификатора"])["Цена по прайсу"] \
    .transform(lambda x: x.fillna(x.mean()))
df_data.sort_values(["Номенклатура.Позиция классификатора", "Дата"], inplace=True)
df_data["Цена по прайсу"] = df_data.groupby("Номенклатура.Позиция классификатора")["Цена по прайсу"].ffill()
df_data["Цена по прайсу"] = df_data["Цена по прайсу"].fillna(df_data["Цена"])
df_data.sort_index(inplace=True)

# Фильтруем строки
df_data = df_data[(df_data['Канал продаж'] == "Прямые продажи") & (df_data['Вид продаж'] == "СЗР")]

# Расчёт сумм для отклонения по прайсу
df_data["Сумма_Факт"] = df_data["Цена"] * df_data["Количество заказано"]
df_data["Сумма_Прайс"] = df_data["Цена по прайсу"] * df_data["Количество заказано"]

# 2. Исход_заказ_с_оплатами
sheet_payments = 'Исход_заказ_с_оплатами'
df_pay = pd.read_excel(path_in, sheet_name=sheet_payments, engine='openpyxl')
cols_to_remove = [
    "Дивизион", "Наименование", "Торговый представитель (ЛИТ)", "Договор", "Сезон",
    "Новый клиент сезона", "Заказ клиента", "Номер", "Вид продаж", "Состояние заказа",
    "График оплаты", "Организация", "% АС1", "Сумма бонуса АС1",
    "УсловиеОплаты", "Заказано", "Отгружено", "Оплачено", "Долг клиента", "Наш долг",
    "Просроченная задолженность", "Дата оплаты (план)", "Дата оплаты (факт)",
    "Дата отгрузки (план)", "Дата отгрузки (факт)", "Количество товаров в заказе"
]
df_pay.drop(columns=cols_to_remove, errors='ignore', inplace=True)
df_pay['Дата'] = pd.to_datetime(df_pay['Дата'], dayfirst=True).dt.date
df_pay['Клиент'] = df_pay['Клиент'].astype(str)
df_pay['Канал продаж'] = df_pay['Канал продаж'].astype(str)
df_pay = df_pay.drop_duplicates()
df_pay = df_pay.sort_values(by=["Клиент", "Дата"], ascending=[True, False])
df_pay_max = df_pay.loc[df_pay.groupby('Клиент')['Дата'].idxmax()].reset_index(drop=True)
df_pay_max = df_pay_max[['Клиент', 'Канал продаж', 'Посевные площади. Га (Общие)', 'Сегмент по площадям (ШАНС)']]
df_pay_max['Посевные площади. Га (Общие)'] = df_pay_max['Посевные площади. Га (Общие)'].round(0).astype('Int64')

# 3. Исход_ДЗ → отгрузки
sheet_dz = 'Исход_ДЗ'
df_dz = pd.read_excel(path_in, sheet_name=sheet_dz, engine='openpyxl')
cols_to_remove_dz = [
    "1", "Организация", "2",
    "Контрагент.Сокращенное юр. наименование", "Контрагент.ИНН",
    "Договор", "Новый клиент сезона", "График оплаты", "Сегмент по площадям (ШАНС)",
    "Посевные площади. Га (Общие)", "Сумма оплаты по заказу, (руб.)",
    "Дата1", "Сумма1", "Дата2", "Сумма2", "Дата3", "Сумма3",
    "Отгружено, (руб.)", "Оплачено, (руб.)",
    "Кредиторская задолженность (КЗ), (руб.)",
    "Нормальная дебиторская задолженность (НДЗ), (руб.)",
    "Просроченная дебиторская задолженность (ПДЗ), (руб.)",
    "Общая дебиторская задолженность, руб", "Дней просрочено",
    "Комментарий договора", "Статус договора", "Оплаты по дням (план)",
    "Оплата по дням (факт)", "Отгрузка по дням (план)"
]
df_dz.drop(columns=cols_to_remove_dz, errors='ignore', inplace=True)
df_dz = df_dz[df_dz["Вид продаж"] == "СЗР"]
df_dz["Регион.Наименование"] = df_dz["Регион.Наименование"].replace({
    "Алтайский край 1": "Алтайский край", "Алтайский край 2": "Алтайский край"
})

def get_division_dz(row):
    region = row["Регион.Наименование"]
    if region in ["Амурская область", "Приморский край"]:
        return "Дивизион ДАЛЬНИЙ ВОСТОК"
    elif region in ["Челябинская область", "Тюменская область", "Курганская область", "Свердловская область"]:
        return "Дивизион УРАЛ"
    elif region in ["Республика Башкортостан", "Республика Татарстан"]:
        return "Дивизион ПОВОЛЖЬЕ"
    elif region in ["Алтайский край", "Кемеровская область", "Красноярский край",
                    "Новосибирская область", "Омская область", "Томская область"]:
        return "Дивизион СИБИРЬ"
    elif region == "Тамбовская область":
        return "Дивизион ЦЕНТР"
    else:
        return row["Дивизион"]

df_dz['Дивизион New'] = df_dz.apply(get_division_dz, axis=1)
df_dz.drop(columns=['Дивизион'], inplace=True)
df_dz.rename(columns={'Дивизион New': 'Дивизион', 'Регион.Наименование': 'Регион'}, inplace=True)
df_dz["Отгрузка по дням (факт)"] = df_dz["Отгрузка по дням (факт)"].astype(str).str.split('\n')
df_dz = df_dz.explode("Отгрузка по дням (факт)")
df_dz = df_dz[df_dz["Отгрузка по дням (факт)"].str.strip() != ""]
df_shipments = df_dz.groupby(["Клиент", "Сезон", "Торговый представитель"])[
    "Отгрузка по дням (факт)"].nunique().reset_index()
df_shipments.rename(columns={"Отгрузка по дням (факт)": "Количество отгрузок"}, inplace=True)

# 4. ABC-анализ: границы
group_cols_boundaries = ["Дивизион", "Регион", "Сезон", "Клиент"]
sales_col = "Сумма заказанной номенклатуры"
df_bound = df_data.groupby(group_cols_boundaries, as_index=False)[sales_col].sum()
df_bound.rename(columns={sales_col: "Продажи, руб."}, inplace=True)
df_bound = df_bound.sort_values(["Дивизион", "Регион", "Сезон", "Продажи, руб."],
                                ascending=[True, True, True, False])
df_bound['CumulativeSales'] = df_bound.groupby(["Дивизион", "Регион", "Сезон"])["Продажи, руб."].cumsum()
df_bound['TotalSales'] = df_bound.groupby(["Дивизион", "Регион", "Сезон"])["Продажи, руб."].transform('sum')
df_bound['CumulativeShare'] = df_bound['CumulativeSales'] / df_bound['TotalSales']
df_bound['ABC_Group'] = df_bound['CumulativeShare'].apply(lambda x: 'A' if x <= 0.80 else ('B' if x <= 0.95 else 'C'))

def assign_category_bound(group):
    group = group.sort_values("Продажи, руб.", ascending=False)
    group.iloc[0, group.columns.get_loc("ABC_Group")] = "A"
    for idx in group.index[1:]:
        cs = group.loc[idx, "CumulativeShare"]
        if cs <= 0.80:
            group.at[idx, "ABC_Group"] = "A"
        elif cs <= 0.95:
            group.at[idx, "ABC_Group"] = "B"
        else:
            group.at[idx, "ABC_Group"] = "C"
    return group

df_bound = df_bound.groupby(["Дивизион", "Регион", "Сезон"], group_keys=False).apply(assign_category_bound)
df_a = (df_bound[df_bound['ABC_Group'] == 'A']
        .groupby(["Дивизион", "Регион", "Сезон"], as_index=False)["Продажи, руб."]
        .min()
        .rename(columns={"Продажи, руб.": "НижняяГраница_A_Премиум"}))
df_b = (df_bound[df_bound['ABC_Group'] == 'B']
        .groupby(["Дивизион", "Регион", "Сезон"], as_index=False)["Продажи, руб."]
        .min()
        .rename(columns={"Продажи, руб.": "НижняяГраница_B_Стандарт"}))
df_boundaries = pd.merge(df_a, df_b, on=["Дивизион", "Регион", "Сезон"], how='outer')

# 5. Распределение по ABC
group_cols_distribution = ["Дивизион", "Регион", "Сезон", "Клиент"]
agg_dict = {
    "Сумма заказанной номенклатуры": "sum",
    "Сумма_Факт": "sum",
    "Сумма_Прайс": "sum"
}
df_dist = df_data.groupby(group_cols_distribution, as_index=False).agg(agg_dict)
df_dist = df_dist.sort_values(["Регион", "Сезон", "Сумма заказанной номенклатуры"],
                              ascending=[True, True, False])
df_dist['CumulativeSales'] = df_dist.groupby(["Дивизион", "Регион", "Сезон"])["Сумма заказанной номенклатуры"].cumsum()
df_dist['TotalSales'] = df_dist.groupby(["Дивизион", "Регион", "Сезон"])["Сумма заказанной номенклатуры"].transform('sum')
df_dist['CumulativeShare'] = df_dist['CumulativeSales'] / df_dist['TotalSales']
df_dist['Категория ABC-анализа'] = df_dist['CumulativeShare'].apply(
    lambda x: 'A' if x <= 0.80 else ('B' if x <= 0.95 else 'C'))
#fdsfsdf
def assign_category_dist(group):
    group = group.sort_values("Сумма заказанной номенклатуры", ascending=False)
    group.iloc[0, group.columns.get_loc("Категория ABC-анализа")] = "A"
    for idx in group.index[1:]:
        cs = group.loc[idx, "CumulativeShare"]
        if cs <= 0.80:
            group.at[idx, "Категория ABC-анализа"] = "A"
        elif cs <= 0.95:
            group.at[idx, "Категория ABC-анализа"] = "B"
        else:
            group.at[idx, "Категория ABC-анализа"] = "C"
    return group

df_dist = df_dist.groupby(["Дивизион", "Регион", "Сезон"], group_keys=False).apply(assign_category_dist)
df_dist["Отклонение_от_прайса%"] = (df_dist["Сумма_Факт"] - df_dist["Сумма_Прайс"]) / df_dist["Сумма_Прайс"]

# Добавляем счётчики, площади и Сегмент по площадям
df_counts = df_data.groupby(group_cols_distribution, as_index=False).agg({
    "Заказ клиента": pd.Series.nunique,
    "Группа аналитического учета": pd.Series.nunique,
    "Номенклатура.Позиция классификатора": pd.Series.nunique
})
df_counts.rename(columns={
    "Заказ клиента": "Количество заказов",
    "Группа аналитического учета": "Кол-во групп товаров",
    "Номенклатура.Позиция классификатора": "Кол-во препаратов"
}, inplace=True)
df_dist = pd.merge(df_dist, df_counts, on=group_cols_distribution, how='left')
df_dist = pd.merge(
    df_dist,
    df_pay_max[['Клиент', 'Посевные площади. Га (Общие)', 'Сегмент по площадям (ШАНС)']],
    on='Клиент', how='left'
)

def _segment_area(v):
    if pd.isna(v):
        return np.nan
    if v >= 40000:
        return "> 40000"
    elif v >= 30000:
        return "30000 - 40000"
    elif v >= 20000:
        return "20000 - 30000"
    elif v >= 10000:
        return "10000 - 20000"
    elif v >= 5000:
        return "5000 - 10000"
    elif v >= 1000:
        return "1000 - 5000"
    elif v >= 500:
        return "500 - 1000"
    else:
        return "< 500"

df_dist["Сегмент по площадям"] = df_dist["Посевные площади. Га (Общие)"].apply(_segment_area)

# Переставим столбец рядом с площадями (для удобства — опционально)
_cols = list(df_dist.columns)
if "Посевные площади. Га (Общие)" in _cols and "Сегмент по площадям" in _cols:
    area_idx = _cols.index("Посевные площади. Га (Общие)")
    seg_idx = _cols.index("Сегмент по площадям")
    _cols.insert(area_idx + 1, _cols.pop(seg_idx))
    df_dist = df_dist[_cols]

# Отгрузки
df_dist = pd.merge(df_dist, df_shipments.groupby(["Клиент", "Сезон"])["Количество отгрузок"].sum().reset_index(),
                   on=['Клиент', 'Сезон'], how='left')

# Детализация по группам товаров
df_group = df_data.groupby(
    ["Дивизион", "Регион", "Сезон", "Клиент", "Группа аналитического учета"],
    as_index=False
)["Номенклатура.Позиция классификатора"].nunique()
df_group.rename(columns={"Номенклатура.Позиция классификатора": "Кол-во препаратов по группе"}, inplace=True)
df_pivot = df_group.pivot_table(
    index=["Дивизион", "Регион", "Сезон", "Клиент"],
    columns="Группа аналитического учета",
    values="Кол-во препаратов по группе",
    fill_value=0
).reset_index()
df_pivot = df_pivot.replace(0, np.nan)
df_distribution = pd.merge(df_dist, df_pivot, on=["Дивизион", "Регион", "Сезон", "Клиент"], how='left')

cols_to_drop = ["Сумма_Факт", "Сумма_Прайс", "CumulativeSales", "TotalSales", "CumulativeShare"]
df_distribution.drop(columns=cols_to_drop, errors='ignore', inplace=True)

df_distribution = df_distribution[df_distribution["Сумма заказанной номенклатуры"] != 0]

# Потенциал, доля, ценовая категория, A+
df_distribution["Примерный объем закупок клиента СЗР"] = df_distribution["Посевные площади. Га (Общие)"].apply(
    lambda x: x if pd.notnull(x) and x != 0 else 500) * 4500
df_distribution["Наша доля"] = df_distribution["Сумма заказанной номенклатуры"] / df_distribution["Примерный объем закупок клиента СЗР"]

def calculate_potential(row):
    abc = row["Категория ABC-анализа"]
    share = row["Наша доля"]
    if abc == "A":
        return "A" if share < 0.20 else "B"
    else:
        return "C" if share < 0.20 else "D"

df_distribution["Категория потенциала"] = df_distribution.apply(calculate_potential, axis=1)

def calc_price_category(x):
    if x >= 0:
        return "A"
    elif x < -0.10:
        return "C -"
    elif x < 0 and x >= -0.05:
        return "B"
    else:
        return "C"

df_distribution["Категория ABC-анализа с отклонением по цене"] = df_distribution["Отклонение_от_прайса%"].apply(calc_price_category)

def split_A(row):
    if row["Категория ABC-анализа"] == "A":
        return "A+" if row["Сумма заказанной номенклатуры"] > 10000000 else "A"
    else:
        return row["Категория ABC-анализа"]

df_distribution["Категория A_A+_B_C-анализа"] = df_distribution.apply(split_A, axis=1)

# Потеря в следующем сезоне
max_season = df_distribution["Сезон"].astype(int).max()
client_seasons = df_distribution.groupby("Клиент")["Сезон"].apply(lambda s: set(pd.to_numeric(s, errors='coerce'))).to_dict()

def lost_next_season(row):
    try:
        current = int(row["Сезон"])
    except:
        return ""
    if current == max_season:
        return "Следующий сезон не наступил"
    next_season = current + 1
    client = row["Клиент"]
    return "Потерян в следующем сезоне" if next_season not in client_seasons.get(client, set()) else ""

df_distribution["Потерянные для следующего сезона"] = df_distribution.apply(lost_next_season, axis=1)

# Расширенные подписи
map_abc = {'A': 'Премиум (A)', 'B': 'Стандарт (B)', 'C': 'Эконом (C)'}
map_pot = {
    'A': 'Драйверы - Берут много и могут больше (А)',
    'B': 'Потолок - Берут много, но больше не могут (В)',
    'C': 'Резервы - Берут мало, но могут больше (С)',
    'D': 'Балласт - Берут мало и больше не могут (D)'
}
df_distribution['Категория ABC-анализа расширенный'] = df_distribution['Категория ABC-анализа'].map(map_abc)
df_distribution['Категория потенциала расширенный'] = df_distribution['Категория потенциала'].map(map_pot)

# Флаг контрактации
df_distribution["Контрактация > 500 000 р."] = df_distribution["Сумма заказанной номенклатуры"].apply(
    lambda x: "Контрактация > 500 000 р." if x >= 500000 else ""
)

# Добавляем столбцы для миграции
df_distribution = df_distribution.sort_values(["Клиент", "Сезон"])
df_distribution["prev_season"] = df_distribution.groupby("Клиент")["Сезон"].shift(1)
df_distribution["prev_abc"] = df_distribution.groupby("Клиент")["Категория ABC-анализа"].shift(1)
df_distribution["Сезоны миграции"] = np.where(
    df_distribution["prev_season"] == df_distribution["Сезон"] - 1,
    df_distribution["Сезон"].astype(str) + " к " + df_distribution["prev_season"].astype(str),
    ""
)
df_distribution["Миграция клиента"] = np.where(
    df_distribution["prev_season"].isna() | (df_distribution["prev_season"] != df_distribution["Сезон"] - 1),
    "Новый клиент",
    df_distribution["prev_abc"] + " - " + df_distribution["Категория ABC-анализа"]
)
df_distribution.drop(columns=["prev_season", "prev_abc"], inplace=True)

# "ABC и ABCD-анализ" = "<Категория ABC-анализа>_<Категория потенциала>"
df_distribution["ABC и ABCD-анализ"] = (
    df_distribution["Категория ABC-анализа"].astype(str)
    + "_"
    + df_distribution["Категория потенциала"].astype(str)
)

# Порядок столбцов
final_columns_dist = [
    "Дивизион", "Регион", "Клиент",
    "Посевные площади. Га (Общие)", "Сегмент по площадям", "Сегмент по площадям (ШАНС)", "Сезон",
    "Сумма заказанной номенклатуры", "Категория ABC-анализа", "Категория ABC-анализа расширенный",
    "Отклонение_от_прайса%", "Количество заказов", "Количество отгрузок", "Кол-во групп товаров",
    "Кол-во препаратов", "Адьюванты", "Гербициды", "Десиканты", "Инсектициды", "Микроудобрения",
    "Протравители", "Регуляторы роста", "Родентициды", "Фумиганты", "Фунгициды",
    "Примерный объем закупок клиента СЗР", "Наша доля", "Категория потенциала",
    "Категория потенциала расширенный", "Категория ABC-анализа с отклонением по цене",
    "Категория A_A+_B_C-анализа", "Потерянные для следующего сезона", "Контрактация > 500 000 р.",
    "Сезоны миграции", "Миграция клиента", "ABC и ABCD-анализ"
]
missing_cols = [col for col in final_columns_dist if col not in df_distribution.columns]
if missing_cols:
    raise ValueError(f"Missing columns in df_distribution: {missing_cols}")
df_distribution = df_distribution[final_columns_dist]

# 6. Миграция по ABC
df_mig = df_distribution[["Дивизион", "Регион", "Клиент", "Сезон", "Категория ABC-анализа", "Категория ABC-анализа расширенный"]].copy()
df_mig["Сезон"] = pd.to_numeric(df_mig["Сезон"], errors='coerce')
df_mig = df_mig.sort_values(["Клиент", "Сезон"])
df_mig["prev_abc"] = df_mig.groupby("Клиент")["Категория ABC-анализа"].shift(1)
df_mig["prev_extended"] = df_mig.groupby("Клиент")["Категория ABC-анализа расширенный"].shift(1)
df_mig["prev_season"] = df_mig.groupby("Клиент")["Сезон"].shift(1)
df_mig = df_mig[(df_mig["Сезон"] - df_mig["prev_season"]) == 1]
df_mig["Сезоны"] = df_mig.apply(lambda r: f"{int(r['Сезон'])} к {int(r['prev_season'])}", axis=1)
df_mig["Миграция"] = "Из " + df_mig["prev_extended"].astype(str) + " в " + df_mig["Категория ABC-анализа расширенный"].astype(str)
df_mig_counts = df_mig.groupby(["Дивизион", "Регион", "Сезоны", "Миграция"], as_index=False)["Клиент"].nunique()
df_mig_pivot = df_mig_counts.pivot_table(index=["Дивизион", "Регион", "Сезоны"], columns="Миграция", values="Клиент", fill_value=0).reset_index()

grouped = df_distribution.groupby(["Дивизион", "Регион"])["Сезон"].unique().reset_index()
grouped["Сезоны_list"] = grouped["Сезон"].apply(lambda x: sorted(x))
rows = []
for _, row in grouped.iterrows():
    div = row["Дивизион"]
    reg = row["Регион"]
    seasons = row["Сезоны_list"]
    if len(seasons) == 0:
        continue
    min_season = min(seasons)
    for s in seasons:
        if s > min_season:
            rows.append({"Дивизион": div, "Регион": reg, "Сезоны": f"{s} к {s-1}", "prev_season": s - 1})
df_transitions = pd.DataFrame(rows)
df_transitions["prev_season"] = df_transitions["prev_season"].astype(int)

df_base = pd.merge(df_transitions, df_mig_pivot, on=["Дивизион", "Регион", "Сезоны"], how="left")
migration_cols = [c for c in df_base.columns if c not in ["Дивизион", "Регион", "Сезоны", "prev_season"]]
df_base[migration_cols] = df_base[migration_cols].fillna(0).astype(int)

df_distribution["Сезон"] = df_distribution["Сезон"].astype(int)
df_A = df_distribution[df_distribution["Категория ABC-анализа"] == "A"].groupby(["Дивизион", "Регион", "Сезон"], as_index=False)["Клиент"].nunique()
df_A.rename(columns={"Клиент": "Кол-во клиентов в категории A на начало сезона"}, inplace=True)
df_B = df_distribution[df_distribution["Категория ABC-анализа"] == "B"].groupby(["Дивизион", "Регион", "Сезон"], as_index=False)["Клиент"].nunique()
df_B.rename(columns={"Клиент": "Кол-во клиентов в категории B на начало сезона"}, inplace=True)
df_C = df_distribution[df_distribution["Категория ABC-анализа"] == "C"].groupby(["Дивизион", "Регион", "Сезон"], as_index=False)["Клиент"].nunique()
df_C.rename(columns={"Клиент": "Кол-во клиентов в категории C на начало сезона"}, inplace=True)

for add_df in (df_A, df_B, df_C):
    df_base = pd.merge(df_base, add_df, left_on=["Дивизион", "Регион", "prev_season"], right_on=["Дивизион", "Регион", "Сезон"], how="left")
    df_base.drop(columns=["Сезон"], inplace=True, errors='ignore')

for col in ["Кол-во клиентов в категории A на начало сезона",
            "Кол-во клиентов в категории B на начало сезона",
            "Кол-во клиентов в категории C на начало сезона"]:
    df_base[col] = df_base[col].fillna(0).astype(int)

df_base.drop(columns=["prev_season"], inplace=True)

# 7. Распределение с ТП
group_cols_tp = ["Дивизион", "Регион", "Сезон", "Клиент", "Торговый представитель"]
agg_dict_tp = {
    "Сумма заказанной номенклатуры": "sum",
    "Сумма_Факт": "sum",
    "Сумма_Прайс": "sum",
    "Заказ клиента": pd.Series.nunique,
    "Группа аналитического учета": pd.Series.nunique,
    "Номенклатура.Позиция классификатора": pd.Series.nunique
}
df_dist_tp = df_data.groupby(group_cols_tp, as_index=False).agg(agg_dict_tp)
df_dist_tp.rename(columns={
    "Заказ клиента": "Количество заказов",
    "Группа аналитического учета": "Кол-во групп товаров",
    "Номенклатура.Позиция классификатора": "Кол-во препаратов"
}, inplace=True)
df_dist_tp["Отклонение_от_прайса%"] = (df_dist_tp["Сумма_Факт"] - df_dist_tp["Сумма_Прайс"]) / df_dist_tp["Сумма_Прайс"]
df_dist_tp.drop(columns=["Сумма_Факт", "Сумма_Прайс"], inplace=True)

df_dist_tp = pd.merge(df_dist_tp, df_shipments, on=["Клиент", "Сезон", "Торговый представитель"], how='left')

df_group_tp = df_data.groupby(
    ["Дивизион", "Регион", "Сезон", "Клиент", "Торговый представитель", "Группа аналитического учета"],
    as_index=False
)["Номенклатура.Позиция классификатора"].nunique()
df_group_tp.rename(columns={"Номенклатура.Позиция классификатора": "Кол-во препаратов по группе"}, inplace=True)
df_pivot_tp = df_group_tp.pivot_table(
    index=["Дивизион", "Регион", "Сезон", "Клиент", "Торговый представитель"],
    columns="Группа аналитического учета",
    values="Кол-во препаратов по группе",
    fill_value=0
).reset_index()
df_pivot_tp = df_pivot_tp.replace(0, np.nan)
df_dist_tp = pd.merge(df_dist_tp, df_pivot_tp, on=["Дивизион", "Регион", "Сезон", "Клиент", "Торговый представитель"], how='left')

cols_from_dist = [
    "Дивизион", "Регион", "Сезон", "Клиент", "Посевные площади. Га (Общие)", "Сегмент по площадям", "Сегмент по площадям (ШАНС)",
    "Категория ABC-анализа", "Категория ABC-анализа расширенный", "Примерный объем закупок клиента СЗР",
    "Наша доля", "Категория потенциала", "Категория потенциала расширенный",
    "Категория A_A+_B_C-анализа", "Потерянные для следующего сезона", "Контрактация > 500 000 р."
]
df_dist_tp = pd.merge(df_dist_tp, df_distribution[cols_from_dist], on=["Дивизион", "Регион", "Сезон", "Клиент"], how='left')

df_dist_tp["Категория ABC-анализа с отклонением по цене"] = df_dist_tp["Отклонение_от_прайса%"].apply(calc_price_category)
df_dist_tp = df_dist_tp[df_dist_tp["Сумма заказанной номенклатуры"] != 0]
df_dist_tp["Наша доля ограниченная"] = df_dist_tp["Наша доля"].apply(lambda x: min(x, 1.0) if pd.notnull(x) else x)

# Ключи
sheet_keys = 'Ключи'
df_keys = pd.read_excel(path_in, sheet_name=sheet_keys, engine='openpyxl')
df_keys = df_keys[['ФИО', 'Должность', 'Стаж в Ко (годы)']]
df_dist_tp = pd.merge(df_dist_tp, df_keys, left_on='Торговый представитель', right_on='ФИО', how='left')
df_dist_tp.drop(columns=['ФИО'], inplace=True)

# 8. Рейтинг ТП

# Баллы по клиенту
shans_to_score = {"A+": 5, "A": 4, "B": 3, "C": 2, "D": 1}
abc_to_score = {"A": 5, "B": 3, "C": 1}
abcd_to_score = {"A": 5, "B": 4, "C": 3, "D": 1}

df_dist_tp["Сегмент по площадям Шанс_Балл"] = df_dist_tp["Сегмент по площадям (ШАНС)"].map(shans_to_score).fillna(2)
df_dist_tp["ABC_балл"] = df_dist_tp["Категория ABC-анализа"].map(abc_to_score).fillna(1)
df_dist_tp["ABCD_балл"] = df_dist_tp["Категория потенциала"].map(abcd_to_score).fillna(1)
df_dist_tp["Ср_балл: Сегмент по площадям_ABC_ABCD"] = (
    df_dist_tp["Сегмент по площадям Шанс_Балл"] + df_dist_tp["ABC_балл"] + df_dist_tp["ABCD_балл"]
) / 3

# Штрафы
potential_mapping = {"A": 4, "B": 3, "C": 2, "D": 1}
df_dist_tp["Категория потенциала числовое"] = df_dist_tp["Категория потенциала"].map(potential_mapping)
df_dist_tp["Потерянные для следующего сезона исходное"] = df_dist_tp["Потерянные для следующего сезона"]
df_dist_tp["Потерянные для следующего сезона числовое"] = df_dist_tp["Потерянные для следующего сезона"].apply(
    lambda x: 1 if x == "Потерян в следующем сезоне" else 0
)
df_dist_tp["Штраф_за_потерю"] = df_dist_tp.apply(
    lambda row: (row["Категория потенциала числовое"] - 1) * row["Потерянные для следующего сезона числовое"],
    axis=1
)
df_dist_tp["Потерянные для следующего сезона"] = df_dist_tp.apply(
    lambda row: "Потерян в следующем сезоне" if row["Потерянные для следующего сезона числовое"] == 1 else row["Потерянные для следующего сезона исходное"],
    axis=1
)
df_dist_tp.drop(columns=["Потерянные для следующего сезона исходное"], inplace=True)

# Застрявшие C
stuck_c_assignments = []
data_sorted = df_dist_tp.sort_values(["Клиент", "Сезон"]).copy()
for client in data_sorted["Клиент"].unique():
    client_data = data_sorted[data_sorted["Клиент"] == client][["Сезон", "Категория ABC-анализа", "Торговый представитель"]].reset_index(drop=True)
    seasons = client_data["Сезон"].values
    categories = client_data["Категория ABC-анализа"].values
    reps = client_data["Торговый представитель"].values
    for i in range(len(seasons) - 2):
        if (categories[i] == "C" and categories[i+1] == "C" and categories[i+2] == "C"
                and seasons[i+1] == seasons[i] + 1 and seasons[i+2] == seasons[i+1] + 1):
            stuck_c_assignments.append({"Клиент": client, "Сезон": seasons[i+2], "Торговый представитель": reps[i+2]})
stuck_c_df = pd.DataFrame(stuck_c_assignments)

def calculate_stuck_penalty(num_stuck_clients):
    if num_stuck_clients >= 5:
        return 2.0
    elif num_stuck_clients >= 3:
        return 1.0
    else:
        return 0

def normalize(series):
    if series.max() == series.min():
        return pd.Series(0, index=series.index)
    return (series - series.min()) / (series.max() - series.min())

# Подготовка результатов
results = []
unique_combinations = df_dist_tp[["Дивизион", "Сезон"]].drop_duplicates()

for _, row in unique_combinations.iterrows():
    division = row["Дивизион"]
    season = row["Сезон"]
    subset = df_dist_tp[(df_dist_tp["Дивизион"] == division) & (df_dist_tp["Сезон"] == season)]
    if subset.empty:
        continue

    grouped = subset.groupby("Торговый представитель")
    client_counts = grouped["Клиент"].nunique()
    max_clients = client_counts.max()

    stuck_counts = stuck_c_df[stuck_c_df["Сезон"] == season].groupby("Торговый представитель")["Клиент"].count().reindex(grouped.groups.keys(), fill_value=0)

    indicators = pd.DataFrame({
        "Total_Sum": grouped["Сумма заказанной номенклатуры"].sum(),
        "Avg_Sum_per_Client": grouped["Сумма заказанной номенклатуры"].mean(),
        "Num_Clients": grouped["Ср_балл: Сегмент по площадям_ABC_ABCD"].mean() * (client_counts / max_clients),
        "Avg_Share": grouped["Наша доля ограниченная"].mean() * (client_counts / max_clients),
        "Penalty_Lost_Clients": grouped["Штраф_за_потерю"].sum(),
        "Product_Diversity": grouped["Кол-во групп товаров"].mean() * (client_counts / max_clients),
        "Num_Stuck_C_Clients": stuck_counts,
        "Penalty_Low_Client_Count": client_counts.apply(lambda x: 1 / x if x > 0 else 1.0)
    })

    indicators["Penalty_Stuck_C_Clients"] = indicators["Num_Stuck_C_Clients"].apply(calculate_stuck_penalty)

    metrics_to_normalize = [
        "Total_Sum", "Avg_Sum_per_Client", "Num_Clients", "Avg_Share",
        "Product_Diversity", "Penalty_Lost_Clients", "Penalty_Stuck_C_Clients",
        "Penalty_Low_Client_Count"
    ]
    for metric in metrics_to_normalize:
        indicators[f"Norm_{metric}"] = normalize(indicators[metric])

    weights = {
        "Norm_Total_Sum": 0.3,
        "Norm_Avg_Sum_per_Client": 0.1,
        "Norm_Num_Clients": 0.25,
        "Norm_Avg_Share": 0.15,
        "Norm_Penalty_Lost_Clients": -0.10,
        "Norm_Product_Diversity": 0.2,
        "Norm_Penalty_Stuck_C_Clients": -0.10,
        "Norm_Penalty_Low_Client_Count": -0.00
    }

    # создаём вклад для каждого нормализованного показателя без ручных имён
    for col, w in weights.items():
        indicators[f"Вклад_{col}"] = indicators[col] * w

    # итоговый скор
    indicators["Score"] = sum(indicators[col] * w for col, w in weights.items())

    indicators["Дивизион"] = division
    indicators["Сезон"] = season
    indicators = indicators.reset_index()  # index -> "Торговый представитель"

    # Формируем список столбцов к выводу динамически
    base_metrics_order = [
        ("Total_Sum",),
        ("Avg_Sum_per_Client",),
        ("Num_Clients",),
        ("Avg_Share",),
        ("Penalty_Lost_Clients",),
        ("Product_Diversity",),
        ("Num_Stuck_C_Clients", "Penalty_Stuck_C_Clients"),  # пара связанных
        ("Penalty_Low_Client_Count",)
    ]

    output_cols = ["Дивизион", "Торговый представитель", "Сезон"]
    for group in base_metrics_order:
        for m in group:
            output_cols.append(m)
            # если есть нормализованная версия — добавим её и вклад
            norm_name = f"Norm_{m}"
            if norm_name in indicators.columns:
                output_cols.append(norm_name)
                contrib_name = f"Вклад_{norm_name}"
                if contrib_name in indicators.columns:
                    output_cols.append(contrib_name)

    output_cols.append("Score")

    # Гарантируем наличие столбцов
    output_cols = [c for c in output_cols if c in indicators.columns]
    df = indicators[output_cols]
    results.append(df)

# Объединение результатов
final_df = pd.concat(results, ignore_index=True)

# Ранг в разрезе дивизиона и сезона
final_df["Место в рейтинге"] = final_df.groupby(["Дивизион", "Сезон"])["Score"].rank(method="min", ascending=False).astype(int)
final_df = final_df.sort_values(by=["Дивизион", "Сезон", "Score"], ascending=[True, True, False])

# Читаемые заголовки
column_names = {
    "Total_Sum": "Общая сумма заказанной номенклатуры",
    "Norm_Total_Sum": "Нормализованная общая сумма",
    "Avg_Sum_per_Client": "Средняя сумма на клиента",
    "Norm_Avg_Sum_per_Client": "Нормализованная средняя сумма на клиента",
    "Num_Clients": "Ср_балл: Сегмент по площадям_ABC_ABCD",
    "Norm_Num_Clients": "Нормализованный ср. балл: Сегмент по площадям_ABC_ABCD",
    "Avg_Share": "Средняя доля в бюджете клиента",
    "Norm_Avg_Share": "Нормализованная средняя доля",
    "Penalty_Lost_Clients": "Штраф за потерянных клиентов",
    "Norm_Penalty_Lost_Clients": "Нормализованный штраф за потерянных клиентов",
    "Product_Diversity": "Разнообразие продуктов",
    "Norm_Product_Diversity": "Нормализованное разнообразие продуктов",
    "Num_Stuck_C_Clients": "Количество застрявших клиентов C",
    "Penalty_Stuck_C_Clients": "Штраф за застрявших клиентов C",
    "Norm_Penalty_Stuck_C_Clients": "Нормализованный штраф за застрявших клиентов C",
    "Penalty_Low_Client_Count": "Штраф за малое количество клиентов",
    "Norm_Penalty_Low_Client_Count": "Нормализованный штраф за малое количество клиентов",
    "Score": "Итоговый балл",
    "Вклад_Norm_Total_Sum": "Вклад общей суммы",
    "Вклад_Norm_Avg_Sum_per_Client": "Вклад средней суммы на клиента",
    "Вклад_Norm_Num_Clients": "Вклад ср балла: Сегмент по площадям_ABC_ABCD",
    "Вклад_Norm_Avg_Share": "Вклад средней доли",
    "Вклад_Norm_Penalty_Lost_Clients": "Вклад штрафа за потерянных клиентов",
    "Вклад_Norm_Product_Diversity": "Вклад разнообразия продуктов",
    "Вклад_Norm_Penalty_Stuck_C_Clients": "Вклад штрафа за застрявших клиентов C",
    "Вклад_Norm_Penalty_Low_Client_Count": "Вклад штрафа за малое количество клиентов",
    "Место в рейтинге": "Место в рейтинге"
}
final_df = final_df.rename(columns=column_names)

# Порядок столбцов для листа с ТП
final_columns_tp = [
    "Дивизион", "Регион", "Клиент", "Торговый представитель", "Должность", "Стаж в Ко (годы)",
    "Посевные площади. Га (Общие)", "Сегмент по площадям", "Сегмент по площадям (ШАНС)", "Сезон", "Сумма заказанной номенклатуры",
    "Категория ABC-анализа", "Категория ABC-анализа расширенный", "Отклонение_от_прайса%",
    "Количество заказов", "Количество отгрузок", "Кол-во групп товаров", "Кол-во препаратов",
    "Адьюванты", "Гербициды", "Десиканты", "Инсектициды", "Микроудобрения", "Протравители",
    "Регуляторы роста", "Родентициды", "Фумиганты", "Фунгициды", "Примерный объем закупок клиента СЗР",
    "Наша доля", "Категория потенциала", "Категория потенциала расширенный",
    "Категория ABC-анализа с отклонением по цене", "Категория A_A+_B_C-анализа",
    "Потерянные для следующего сезона", "Сегмент по площадям Шанс_Балл",
    "Ср_балл: Сегмент по площадям_ABC_ABCD", "Контрактация > 500 000 р."
]
df_dist_tp = df_dist_tp[final_columns_tp]

# 9. Клиенты по сезонам
cols_for_pivot = [
    "Дивизион", "Регион", "Клиент", "Сезон", "Посевные площади. Га (Общие)", "Сегмент по площадям", "Сегмент по площадям (ШАНС)",
    "Сумма заказанной номенклатуры", "Категория ABC-анализа расширенный",
    "Категория потенциала расширенный", "Кол-во групп товаров", "Кол-во препаратов"
]
df_clients_by_season = df_distribution[cols_for_pivot].copy()
df_clients_by_season["Сезон"] = df_clients_by_season["Сезон"].astype(str)

all_clients = df_data["Клиент"].unique()
clients_in_season = df_clients_by_season["Клиент"].unique()
missing_clients = set(all_clients) - set(clients_in_season)
if missing_clients:
    for client in missing_clients:
        df_missing = pd.DataFrame({
            "Дивизион": [np.nan], "Регион": [np.nan], "Клиент": [client], "Сезон": [np.nan],
            "Посевные площади. Га (Общие)": [np.nan], "Сегмент по площадям": [np.nan], "Сегмент по площадям (ШАНС)": [np.nan],
            "Сумма заказанной номенклатуры": [np.nan], "Категория ABC-анализа расширенный": [np.nan],
            "Категория потенциала расширенный": [np.nan], "Кол-во групп товаров": [np.nan],
            "Кол-во препаратов": [np.nan]
        })
        df_clients_by_season = pd.concat([df_clients_by_season, df_missing], ignore_index=True)

df_clients_by_season.fillna(value={
    "Дивизион": "Не определен",
    "Регион": "Не определен",
    "Сезон": "Не определен",
    "Посевные площади. Га (Общие)": 0,
    "Сегмент по площадям": "Не определен",
    "Сегмент по площадям (ШАНС)": "Не определен",
    "Сумма заказанной номенклатуры": 0,
    "Категория ABC-анализа расширенный": "Не определен",
    "Категория потенциала расширенный": "Не определен",
    "Кол-во групп товаров": 0,
    "Кол-во препаратов": 0
}, inplace=True)

metrics = [
    "Сумма заказанной номенклатуры",
    "Категория ABC-анализа расширенный",
    "Категория потенциала расширенный",
    "Кол-во групп товаров",
    "Кол-во препаратов"
]

pivot_dfs = []
for metric in metrics:
    pivot = df_clients_by_season.pivot_table(
        index=["Дивизион", "Регион", "Клиент", "Посевные площади. Га (Общие)", "Сегмент по площадям", "Сегмент по площадям (ШАНС)"],
        columns="Сезон",
        values=metric,
        aggfunc="first",
        fill_value=np.nan
    ).reset_index()
    pivot.columns = [f"{metric}_{col}" if col in ["23", "24", "25"] else col for col in pivot.columns]
    pivot_dfs.append(pivot)

df_merged = pivot_dfs[0]
for df in pivot_dfs[1:]:
    df_merged = df_merged.merge(
        df.drop(columns=["Посевные площади. Га (Общие)", "Сегмент по площадям", "Сегмент по площадям (ШАНС)"]),
        on=["Дивизион", "Регион", "Клиент"],
        how="outer",
        suffixes=('', '_dup')
    )
    for col in df_merged.columns:
        if col.endswith('_dup'):
            df_merged.drop(columns=col, inplace=True)

# Счётчики раз по категориям
df_count_c = df_clients_by_season[df_clients_by_season["Категория ABC-анализа расширенный"] == "Эконом (C)"]
df_count_c = df_count_c.groupby(["Дивизион", "Регион", "Клиент"])["Категория ABC-анализа расширенный"].count().reset_index()
df_count_c.rename(columns={"Категория ABC-анализа расширенный": "Количество раз в Эконом (C)"}, inplace=True)
df_merged = df_merged.merge(df_count_c, on=["Дивизион", "Регион", "Клиент"], how="left")
df_merged["Количество раз в Эконом (C)"] = df_merged["Количество раз в Эконом (C)"].fillna(0).astype(int)

df_count_abcd_d = df_clients_by_season[df_clients_by_season["Категория потенциала расширенный"] == "Балласт - Берут мало и больше не могут (D)"]
df_count_abcd_d = df_count_abcd_d.groupby(["Дивизион", "Регион", "Клиент"])["Категория потенциала расширенный"].count().reset_index()
df_count_abcd_d.rename(columns={"Категория потенциала расширенный": "Количество раз в Балласт - Берут мало и больше не могут (D)"}, inplace=True)
df_merged = df_merged.merge(df_count_abcd_d, on=["Дивизион", "Регион", "Клиент"], how="left")
df_merged["Количество раз в Балласт - Берут мало и больше не могут (D)"] = df_merged["Количество раз в Балласт - Берут мало и больше не могут (D)"].fillna(0).astype(int)

final_columns_clients = [
    "Дивизион", "Регион", "Клиент", "Посевные площади. Га (Общие)", "Сегмент по площадям", "Сегмент по площадям (ШАНС)",
    "Сумма заказанной номенклатуры_23", "Сумма заказанной номенклатуры_24", "Сумма заказанной номенклатуры_25",
    "Категория ABC-анализа расширенный_23", "Категория ABC-анализа расширенный_24", "Категория ABC-анализа расширенный_25",
    "Количество раз в Эконом (C)", "Категория потенциала расширенный_23", "Категория потенциала расширенный_24",
    "Категория потенциала расширенный_25", "Количество раз в Балласт - Берут мало и больше не могут (D)",
    "Кол-во групп товаров_23", "Кол-во групп товаров_24", "Кол-во групп товаров_25",
    "Кол-во препаратов_23", "Кол-во препаратов_24", "Кол-во препаратов_25"
]
for col in final_columns_clients:
    if col not in df_merged.columns:
        df_merged[col] = np.nan
df_clients_by_season_final = df_merged[final_columns_clients].sort_values(by=["Дивизион", "Регион", "Клиент"])

# Запись в Excel
with pd.ExcelWriter(path_out, engine='openpyxl') as writer:
    df_boundaries.to_excel(writer, sheet_name='Границы ABC по регионам', index=False)
    df_distribution.to_excel(writer, sheet_name='Распределение кл-в по ABC', index=False)  # здесь добавлен "Сегмент по площадям"
    df_base.to_excel(writer, sheet_name='Миграция по ABC', index=False)
    # final_df уже с читаемыми заголовками; порядок столбцов оставляю как есть
    final_df.to_excel(writer, sheet_name='Рейтинг', index=False)
    df_dist_tp.to_excel(writer, sheet_name='Распределение кл-в по ABC с ТП', index=False)
    df_clients_by_season_final.to_excel(writer, sheet_name='Клиенты по сезонам', index=False)
    # короткая таблица метрик — без изменений
    metrics_table = pd.DataFrame([
        {"Метрика": "Total_Sum", "Описание": "Общая сумма заказанной номенклатуры", "Расчет": "Сумма значений столбца 'Сумма заказанной номенклатуры' по всем клиентам представителя в сезоне", "Вес": 0.3},
        {"Метрика": "Avg_Sum_per_Client", "Описание": "Средняя сумма заказанной номенклатуры на клиента", "Расчет": "Среднее значение столбца 'Сумма заказанной номенклатуры' по клиентам представителя в сезоне", "Вес": 0.1},
        {"Метрика": "Num_Clients", "Описание": "Средний балл по комбинации показателей с учетом количества клиентов", "Расчет": "Средний балл по трём показателям (Сегмент по площадям, ABC-анализ, ABCD-анализ), умноженный на коэффициент (количество клиентов / максимум в группе). Баллы: ШАНС A+=5,A=4,B=3,C=2,D=1; ABC: A=5,B=3,C=1; ABCD: A=5,B=4,C=3,D=1", "Вес": 0.25},
        {"Метрика": "Avg_Share", "Описание": "Средняя доля в бюджете клиента (с учётом числа клиентов)", "Расчет": "Среднее 'Наша доля' с отсечкой >1 до 1, умноженное на (кол-во клиентов / максимум в группе)", "Вес": 0.15},
        {"Метрика": "Penalty_Lost_Clients", "Описание": "Штраф за потерянных клиентов", "Расчет": "Сумма штрафов: (Категория потенциала - 1), где A=4 (штраф 3), B=3 (штраф 2), C=2 (штраф 1), D=1 (штраф 0)", "Вес": -0.10},
        {"Метрика": "Product_Diversity", "Описание": "Разнообразие продаж", "Расчет": "Среднее 'Кол-во групп товаров' * (кол-во клиентов / максимум в группе)", "Вес": 0.2},
        {"Метрика": "Penalty_Stuck_C_Clients", "Описание": "Штраф за клиентов C 3 сезона подряд", "Расчет": "≥5 клиентов: 2.0; 3–4 клиента: 1.0; 0–2 клиента: 0", "Вес": -0.10},
        {"Метрика": "Penalty_Low_Client_Count", "Описание": "Штраф за малое количество клиентов", "Расчет": "Нормализация (1 / количество клиентов)", "Вес": -0.00},
    ])
    metrics_table.to_excel(writer, sheet_name='Метрики', index=False)

print("Готово! Итоговый файл сохранён:", path_out)
print("Время работы кода: {:.2f} секунд".format(time.time() - start_time))