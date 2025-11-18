import os
import pandas as pd
import numpy as np

# 1) Пути к файлам
input_path = r"\\192.168.1.211\управление маркетинга\Аналитический центр\Отчёты\Коммерческая политика\2026\тест.xlsx"
output_path = os.path.join(os.path.dirname(input_path), "Расчет нового бонуса 24 сезон.xlsx")

# 2) Загрузка и очистка названий колонок
orders = pd.read_excel(input_path, sheet_name='data_заказ', parse_dates=['Дата'])
prod = pd.read_excel(input_path, sheet_name='data_товар')
plans_raw = pd.read_excel(input_path, sheet_name='Планы скорректированные в минус', skiprows=1)
for df in (orders, prod, plans_raw):
    df.columns = df.columns.str.strip()

# 3) Ключевые колонки в prod
qty_ordered   = next(c for c in prod.columns if 'количество заказано' in c.lower())
qty_shipped   = next(c for c in prod.columns if 'количество отгружено' in c.lower())
price_list_col = next(c for c in prod.columns if 'цена по прайсу' in c.lower())
price_actual_col = next(c for c in prod.columns if c.lower() == 'цена')

# 4) План-лист 24 сезона
c_rep  = next(c for c in plans_raw.columns if 'Торговый представитель' in c)
c_chan = next(c for c in plans_raw.columns if 'Канал' in c)
c_seas = next(c for c in plans_raw.columns if 'Сезон' in c)
c_plan = next(c for c in plans_raw.columns if 'Итого' in c or 'План' in c)

plans24 = (
    plans_raw
    .rename(columns={c_rep: 'rep', c_chan: 'chan', c_seas: 'season', c_plan: 'plan_l'})
    .query("season==24 and chan=='Прямые продажи'")[['rep', 'plan_l']]
)
plan_dict = plans24.groupby('rep')['plan_l'].sum().to_dict()

# 5) Подготовка справочников prod23/prod24
prod23 = prod[prod['Сезон'] == 23].copy()
prod23[qty_shipped]       = prod23[qty_shipped].fillna(0)
prod23['Отгружено руб']   = prod23['Отгружено руб'].fillna(0).round(2)

prod24 = prod[prod['Сезон'] == 24].copy()
prod24['mrc']    = prod24[price_list_col]
prod24['rc']     = (prod24[price_list_col] * 1.10).round(2)
prod24['order_sum'] = (prod24[price_actual_col] * prod24[qty_ordered]).round(2)
prod24[qty_shipped]     = prod24[qty_shipped].fillna(0)
prod24['Отгружено руб'] = prod24['Отгружено руб'].fillna(0).round(2)

# 6) Фильтрация заказов: прямые продажи, СЗР
filter23 = "`Сезон`==23 and `Канал продаж`=='Прямые продажи' and `Вид продаж`=='СЗР'"
filter24 = "`Сезон`==24 and `Канал продаж`=='Прямые продажи' and `Вид продаж`=='СЗР'"
orders23 = orders.query(filter23).copy()
orders24 = orders.query(filter24).copy()
for df in (orders23, orders24):
    for col in ['Заказано', 'Отгружено', 'Оплачено']:
        df[col] = df[col].fillna(0).round(2)
    df['% АС1'] = df['% АС1'].fillna(0)
    df["Сумма бонуса АС1"] = df["Сумма бонуса АС1"].fillna(0).round(2)

# 7) Merge и расчёт net_amount
m23_all = orders23.merge(
    prod23[['Заказ клиента', 'Номер', qty_shipped, 'Отгружено руб']],
    on=['Заказ клиента', 'Номер'], how='left'
)
m23_all['net_amount'] = m23_all['Отгружено руб'] * (
    1 - (
        m23_all["Сумма бонуса АС1"] /
        m23_all.groupby('Номер')['Отгружено руб']
               .transform('sum')
               .replace(0, np.nan)
    ).fillna(0)
)

m24_all = orders24.merge(
    prod24[[
        'Заказ клиента', 'Номер', 'Группа аналитического учета',
        'Номенклатура.Позиция классификатора', price_actual_col,
        'mrc', 'rc', 'order_sum', qty_shipped, 'Отгружено руб'
    ]],
    on=['Заказ клиента', 'Номер'], how='left'
)
m24_all['net_amount'] = m24_all['Отгружено руб'] * (
    1 - (
        m24_all["Сумма бонуса АС1"] /
        m24_all.groupby('Номер')['Отгружено руб']
               .transform('sum')
               .replace(0, np.nan)
    ).fillna(0)
)

# 8) Отбор валидных заказов по статусу "Закрыт"
valid23_nums = orders23.loc[orders23['Состояние заказа'] == 'Закрыт', 'Номер'].unique()
valid24_nums = orders24.loc[orders24['Состояние заказа'] == 'Закрыт', 'Номер'].unique()
valid23 = m23_all[m23_all['Номер'].isin(valid23_nums)].copy()
valid24 = m24_all[m24_all['Номер'].isin(valid24_nums)].copy()

# 9) Вспомогательные «сырые» метрики
def raw_k3(rep):
    return (
        m24_all[m24_all['Торговый представитель (ЛИТ)'] == rep]
        .groupby('Клиент')['Группа аналитического учета']
        .nunique()
        .mean()
    )

def raw_k4_pct(rep):
    prev = (
        m23_all[m23_all['Торговый представитель (ЛИТ)'] == rep]
        .groupby('Клиент')['Отгружено руб']
        .sum()
    )
    prev_c = {c for c, s in prev.items() if s >= 500_000}
    curr_c = set(m24_all['Клиент'])
    if not prev_c:
        return 100.0
    return len(prev_c & curr_c) / len(prev_c) * 100

def raw_k5(rep):
    gross23 = m23_all[m23_all['Торговый представитель (ЛИТ)'] == rep]['Отгружено руб'].sum()
    gross24 = m24_all[m24_all['Торговый представитель (ЛИТ)'] == rep]['Отгружено руб'].sum()
    if gross23 <= 0:
        return 0.0
    return (gross24 - gross23) / gross23 * 100

# 10) Функции расчёта коэффициентов
def rate(nl, mrc, rc, extra):
    if nl < mrc:
        return 0.012 if extra else 0.01
    if nl <= rc:
        return 0.014 if extra else 0.012
    return 0.015 if extra else 0.013

def calc_k1(fact_l, plan_l):
    if plan_l <= 0:
        return 0.0
    pct = fact_l / plan_l * 100
    if pct < 70:
        return 0.0
    return min(pct / 100, 1.5)

def calc_k2(df, plan_l):
    total_net = df['net_amount'].sum()
    total_l   = df[qty_shipped].sum()
    if total_net <= 0 or total_l <= 0:
        return 0.0
    rem = min(plan_l, total_l)
    bb = be = 0.0
    for _, r in df.sort_values('Дата').iterrows():
        l  = r[qty_shipped]
        nl = r['net_amount'] / l if l > 0 else 0
        in_b = min(l, rem)
        in_e = l - in_b
        rem -= in_b
        bb += in_b * nl * rate(nl, r['mrc'], r['rc'], False)
        be += in_e * nl * rate(nl, r['mrc'], r['rc'], True)
    return (bb + be) / total_net

def calc_k3(rep):
    v = raw_k3(rep)
    if v < 2:   return 0.85
    if v < 3:   return 0.9
    if v < 4:   return 1.0
    return 1.2

def calc_k4(rep):
    pct = raw_k4_pct(rep)
    if pct < 65:    return 0.85
    if pct <= 75:   return 1.0
    return 1.1

def calc_k5(rep):
    v = raw_k5(rep)
    if v < 0:    return 0.9
    if v <= 10:  return 1.0
    return 1.1

def calc_k6(rep):
    s23 = (
        m23_all[m23_all['Торговый представитель (ЛИТ)'] == rep]
        .groupby('Клиент')['Отгружено руб']
        .sum()
    )
    s24 = (
        m24_all[m24_all['Торговый представитель (ЛИТ)'] == rep]
        .groupby('Клиент')['Отгружено руб']
        .sum()
    )
    # общие клиенты
    common = set(s23.index) & set(s24.index)
    # новые крупные клиенты
    special = {c for c in s24.index if c not in s23.index and s24[c] >= 500_000}
    denom = common | special
    if not denom:
        return 1.0, 0.0

    small = {c for c in common if s24[c] < 500_000}
    pct_small = len(small) / len(denom) * 100

    if pct_small > 50:
        coeff = 0.7
    elif pct_small >= 40:
        coeff = 0.8
    elif pct_small >= 30:
        coeff = 0.9
    elif pct_small >= 20:
        coeff = 1.0
    else:
        coeff = 1.2

    return coeff, pct_small

def breakdown_k2(df, rep, plan_l, k1, k2, k3, k4, k5, k6):
    rows = []
    rem = min(plan_l, df[qty_shipped].sum())
    denom_idx = df.groupby('Клиент')['Отгружено руб'].sum().index
    numer = {c for c in denom_idx if df.groupby('Клиент')['Отгружено руб'].sum()[c] < 500_000}
    denom_str = ",".join(sorted(denom_idx))
    numer_str = ",".join(sorted(numer))
    for _, r in df.sort_values('Дата').iterrows():
        l = r[qty_shipped]
        if l <= 0: continue
        nl = r['net_amount'] / l
        b = min(l, rem)
        e = l - b
        rem -= b
        lot2 = ((b * rate(nl, r['mrc'], r['rc'], False)) + (e * rate(nl, r['mrc'], r['rc'], True))) / l
        rows.append({
            'РП': rep, 'Клиент': r['Клиент'], 'Номер': r['Номер'],
            'Номенклатура': r['Номенклатура.Позиция классификатора'], 'Дата': r['Дата'],
            'Отгружено, л': l, 'Отгружено руб': r['Отгружено руб'],
            'net_amount': r['net_amount'], 'Цена за л': nl,
            'MRC': r['mrc'], 'RC': r['rc'],
            'lot_k1': k1, 'lot_k2': lot2, 'lot_k3': k3,
            'lot_k4': k4, 'lot_k5': k5, 'lot_k6': k6,
            'denominator_k6': denom_str, 'numerator_k6': numer_str
        })
    return pd.DataFrame(rows)

# 11) Сбор итогов и запись
rows, details = [], []
for rep, grp in valid24.groupby('Торговый представитель (ЛИТ)'):
    plan_l = plan_dict.get(rep, 0.0)
    fact_l = grp[qty_shipped].sum()
    op_net = grp['net_amount'].sum()
    g23    = m23_all[m23_all['Торговый представитель (ЛИТ)'] == rep]['Отгружено руб'].sum()
    g24    = m24_all[m24_all['Торговый представитель (ЛИТ)'] == rep]['Отгружено руб'].sum()
    pct_plan = fact_l / plan_l * 100 if plan_l > 0 else 0.0

    prev = (
        m23_all[m23_all['Торговый представитель (ЛИТ)'] == rep]
        .groupby('Клиент')['Отгружено руб']
        .sum()
    )
    prev_c = {c for c, s in prev.items() if s >= 500_000}
    curr_c = set(m24_all['Клиент'])

    num_K4   = len(prev_c & curr_c)
    denom_K4 = len(prev_c)

    k1 = calc_k1(fact_l, plan_l)
    k2 = calc_k2(grp, plan_l)
    k3 = calc_k3(rep)
    k4 = calc_k4(rep)
    k5 = calc_k5(rep)
    k6, pct_small = calc_k6(rep)

    b1 = 0
    if k1 >= 0.9 and k5 == 1.1 and k6 == 1.2:
        b1 = 100_000
    bonus = op_net * k1 * k2 * k3 * k4 * k5 * k6 + b1

    rows.append({
        'РП': rep, 'OП,руб': op_net, 'gross23': g23, 'gross24': g24,
        'План,л': plan_l, 'Факт,л': fact_l,
        'K1': k1, 'K2': k2, 'K3': k3, 'K4': k4, 'K5': k5, 'K6': k6, 'Б_1': b1,
        'Бонус,руб': bonus, 'pct_small': pct_small,
        'val_K3': raw_k3(rep), 'val_K4': raw_k4_pct(rep),
        'val_K5': raw_k5(rep), 'pct_plan': pct_plan,
        'num_K4': num_K4, 'denom_K4': denom_K4
    })
    details.append(breakdown_k2(grp, rep, plan_l, k1, k2, k3, k4, k5, k6))

results_df = pd.DataFrame(rows)
detail_df  = pd.concat(details, ignore_index=True)

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    results_df.to_excel(writer, sheet_name='Results', index=False)
    detail_df.to_excel(writer, sheet_name='Детали', index=False)

print("✅ Файл сохранён:", output_path)
