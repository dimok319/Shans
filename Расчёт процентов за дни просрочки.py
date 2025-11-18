# Расчёт процентов по просрочкам
# "Проценты, руб." — excel-логика:
#   1) FIFO-распределяем ВСЕ факты по планам (предоплаты разрешены, но платёж не "перепрыгивает" незакрытый план).
#   2) Для каждой плановой строки:
#        - считаем предоплаты (< plan_date), получаем остаток на дату плана;
#        - для КАЖДОГО платежа с датой >= plan_date добавляем проценты:
#              base * (pay_date - plan_date) * (INTEREST_RATE/365),
#          где base:
#              * для ПЕРВОГО события начиная со 2-й плановой строки — "хвост" платежа,
#                перетёкший из закрытия предыдущего плана (carry);
#              * во всех прочих случаях — текущий остаток на момент события;
#        - затем уменьшаем остаток на applied;
#        - если после всех событий остаток > 0 — проценты до today.
# "Проценты на текущий момент, руб." — старая логика: проценты только на текущие остатки (факты ≤ today).
# Заголовки — русские.

import pandas as pd
import re
from datetime import datetime, date

INTEREST_RATE = 0.18  # годовая ставка

# ---------- Утилиты ----------

def clean_amount_string(s: str) -> str:
    s = str(s).replace("\u00A0", "").replace("\u2007", "").replace("\u202F", "")
    s = re.sub(r"[^0-9\.,-]+", "", s)
    return s.replace(",", ".")

def parse_payment_lines(cell):
    recs = []
    if pd.isna(cell) or str(cell).strip() == "":
        return recs
    for line in str(cell).splitlines():
        parts = re.split(r"\s*;\s*", line)
        if len(parts) < 2:
            continue
        try:
            d = datetime.strptime(parts[0].strip(), "%d.%m.%Y").date()
            a = float(clean_amount_string(parts[1].strip()))
            recs.append([d, a])
        except:
            continue
    return sorted(recs, key=lambda x: x[0])

def parse_payment_info(cell, mode="plan", filter_past=False, today=None):
    if pd.isna(cell) or str(cell).strip() == "":
        return 0, None
    recs = []
    for line in str(cell).splitlines():
        parts = re.split(r"\s*;\s*", line)
        if len(parts) < 2:
            continue
        try:
            d = datetime.strptime(parts[0].strip(), "%d.%m.%Y").date()
            a = float(clean_amount_string(parts[1].strip()))
            recs.append((d, a))
        except:
            continue
    if mode == "plan" and filter_past:
        valid = [(d, a) for d, a in recs if d <= today]
        return sum(a for d, a in valid), min((d for d, a in valid), default=None)
    if mode == "plan":
        return sum(a for d, a in recs), min((d for d, a in recs), default=None)
    return sum(a for d, a in recs), max((d for d, a in recs), default=None)

def find_column_by_keywords(df, keywords):
    kws = [k.lower().strip() for k in keywords]
    m = [c for c in df.columns if all(w in c.lower() for w in kws)]
    if not m:
        raise KeyError(f"Столбец с ключами {keywords} не найден.")
    return m[0]

# ---------- FIFO-распределение фактов по планам ----------

def allocate_fifo(plans, facts):
    """
    plans: [(plan_date, plan_amount), ...]   — по возрастанию дат
    facts: [(fact_date, fact_amount), ...]   — по возрастанию дат
    return: allocations — список по планам,
            allocations[i] = [(fact_date, applied_amount), ...]
    Предоплата разрешена, но платёж не "перепрыгивает" незакрытый план.
    """
    allocations = [[] for _ in plans]
    i = 0
    rem = [float(a) for _, a in plans]
    for fd, fa in facts:
        amt = float(fa)
        while amt > 0 and i < len(plans):
            if rem[i] <= 1e-9:
                i += 1
                continue
            applied = min(rem[i], amt)
            allocations[i].append((fd, applied))
            rem[i] -= applied
            amt -= applied
            if rem[i] <= 1e-9:
                i += 1
    return allocations

# ---------- "Проценты, руб." — excel-логика ----------

def calculate_interest_excel_style(row, today):
    plans = parse_payment_lines(row[col_plan])   # [(pdate, pamt), ...]
    facts = parse_payment_lines(row[col_fact])   # [(fdate, famt), ...]
    if not plans:
        return 0.0

    alloc = allocate_fifo(plans, facts)
    r_day = INTEREST_RATE / 365.0
    total = 0.0

    for idx, ((pdate, pamt), a_list) in enumerate(zip(plans, alloc)):
        if pamt <= 0:
            continue

        # предоплаты и остаток на дату плана
        pre  = [(fd, a) for (fd, a) in a_list if fd < pdate]
        post = sorted([(fd, a) for (fd, a) in a_list if fd >= pdate], key=lambda x: x[0])
        pre_total = sum(a for _, a in pre)
        carry     = pre[-1][1] if pre else 0.0     # хвост последнего платежа ДО даты плана (если был)
        remaining = max(0.0, pamt - pre_total)     # остаток на дату плана

        first = True
        for fd, applied in post:
            days = (fd - pdate).days
            if days > 0:
                # ВАЖНО: carry применяем ТОЛЬКО начиная со 2-й плановой строки (idx > 0).
                base = carry if (first and carry > 0 and idx > 0) else remaining
                total += base * r_day * days
            remaining -= applied
            first = False

        # остаток до today
        if remaining > 1e-9 and pdate < today:
            days = (today - pdate).days
            if days > 0:
                total += remaining * r_day * days

    return round(total, 2)

# ---------- "Проценты на текущий момент, руб." — старая логика ----------

def calculate_current_overdue_interest(row, today):
    plans = parse_payment_lines(row[col_plan])
    facts_all = parse_payment_lines(row[col_fact])
    facts = [[d, a] for d, a in facts_all if d <= today]

    # FIFO с предоплатами
    rem = [float(a) for _, a in plans]
    i = 0
    for fd, fa in facts:
        amt = float(fa)
        while amt > 0 and i < len(rem):
            applied = min(rem[i], amt)
            rem[i] -= applied
            amt -= applied
            if rem[i] <= 1e-9:
                i += 1

    r_day = INTEREST_RATE / 365.0
    cur = 0.0
    for (pd0, _), remain in zip(plans, rem):
        if remain > 1e-9 and pd0 < today:
            cur += remain * r_day * (today - pd0).days
    return round(cur, 2)

# ---------- MAIN ----------

if __name__ == "__main__":
    today = date.today()

    input_file = r"\\192.168.1.211\дебиторская задолженность\Отчеты\Выгрузка 1С\ДЗ_1С_НОВЫЙ (XLSX).xlsx"
    output_file = r"C:\Users\nkazakov\Downloads\Расчёт процентов за дни просрочки1.xlsx"

    df = pd.read_excel(input_file, header=2)
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.duplicated()]

    df["Заказ"] = df["Заказ клиента"].astype(str).str.strip()
    col_plan = find_column_by_keywords(df, ["по дням", "план"])
    col_fact = find_column_by_keywords(df, ["по дням", "факт"])

    # Служебные поля
    df[["PlanAmountPast", "PlanDatePast"]] = df[col_plan].apply(
        lambda x: pd.Series(parse_payment_info(x, "plan", True, today))
    )
    df[["PlanAmountFull", "PlanDateFull"]] = df[col_plan].apply(
        lambda x: pd.Series(parse_payment_info(x, "plan", False))
    )
    df[["FactAmountTotal", "FactDateMax"]] = df[col_fact].apply(
        lambda x: pd.Series(parse_payment_info(x, "fact"))
    )

    df["Долг"] = (df["PlanAmountPast"] - df["FactAmountTotal"]).clip(lower=0)
    df["Оплата по дням (факт).руб."] = df["FactAmountTotal"]

    # Привязка регионов к дивизионам (как у вас)
    region_to_division = {
        "Азербайджан": "СНГ",
        "Алтайский край": "Дивизион СИБИРЬ",
        "Амурская область": "Дивизион ДАЛЬНИЙ ВОСТОК",
        "Астраханская область": "Дивизион ЮГ",
        "Белгородская область": "Дивизион ЦЕНТР",
        "Белоруссия": "СНГ",
        "Беларусь": "СНГ",
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
        "Краснодарский край": "Дивизион ЮГ",
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
        "Ростовская область": "Дивизион ЮГ",
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
        "Армения": "СНГ",
        "Алтайский край 1": "Дивизион СИБИРЬ",
        "Алтайский край 2": "Дивизион СИБИРЬ"
    }
    df["Дивизион"] = df["Регион.Наименование"].map(region_to_division).fillna("Неизвестно")

    # Агрегация по заказам
    grouped = df.groupby("Заказ", as_index=False).agg({
        "PlanAmountPast": "sum",
        "FactAmountTotal": "sum",
        "PlanDatePast": "min",
        "FactDateMax": "max",
        "PlanDateFull": "max",
        "Дивизион": "first",
        "Регион.Наименование": "first",
        "Сезон": "first",
        "Контрагент.Сокращенное юр. наименование": "first",
        "Контрагент.ИНН": "first",
        "Канал продаж": "first",
        "Вид продаж": "first",
        "Сумма оплаты по заказу, (руб.)": "first",
        "Оплата по дням (факт).руб.": "first"
    })

    grouped["Агрегированный долг"] = (grouped["PlanAmountPast"] - grouped["FactAmountTotal"]).clip(lower=0)

    # сырьё для процентов
    grouped = grouped.merge(df[["Заказ", col_plan, col_fact]], on="Заказ", how="left")

    # 1) Исторические проценты — excel-логика (исправленный carry)
    grouped["Проценты"] = grouped.apply(lambda r: calculate_interest_excel_style(r, today), axis=1)

    # 2) Проценты на текущий момент — старая логика
    grouped["Проценты на текущий момент"] = grouped.apply(
        lambda r: calculate_current_overdue_interest(r, today), axis=1
    )

    # Финальные столбцы + русские заголовки
    final_columns = [
        "Дивизион", "Регион.Наименование", "Сезон",
        "Контрагент.Сокращенное юр. наименование", "Контрагент.ИНН",
        "Канал продаж", "Вид продаж",
        "Сумма оплаты по заказу, (руб.)", "Оплата по дням (факт).руб.",
        "PlanAmountPast", "FactAmountTotal",
        "PlanDatePast", "FactDateMax", "PlanDateFull",
        "Агрегированный долг", "Проценты", "Проценты на текущий момент",
        "Заказ", col_plan, col_fact
    ]
    grouped = grouped[final_columns].rename(columns={
        "Регион.Наименование": "Регион",
        "Контрагент.Сокращенное юр. наименование": "Контрагент",
        "Контрагент.ИНН": "ИНН",
        "Сумма оплаты по заказу, (руб.)": "Сумма оплаты по заказу, руб.",
        "Оплата по дням (факт).руб.": "Оплата по дням (факт), руб.",
        "PlanAmountPast": "План до сегодня, руб.",
        "FactAmountTotal": "Оплачено факт, руб.",
        "PlanDatePast": "Первая плановая дата ≤ сегодня",
        "FactDateMax": "Последняя дата оплаты (факт)",
        "PlanDateFull": "Первая дата по плану (всего)",
        "Агрегированный долг": "Агрегированный долг, руб.",
        "Проценты": "Проценты, руб.",
        "Проценты на текущий момент": "Проценты на текущий момент, руб.",
        col_plan: "Оплаты по дням (план) — иск",
        col_fact: "Оплаты по дням (факт) — иск"
    })

    for c in [
        "Сумма оплаты по заказу, руб.",
        "Оплата по дням (факт), руб.",
        "План до сегодня, руб.",
        "Оплачено факт, руб.",
        "Агрегированный долг, руб.",
        "Проценты, руб.",
        "Проценты на текущий момент, руб.",
    ]:
        if c in grouped.columns:
            grouped[c] = grouped[c].round(2)

    grouped.to_excel(output_file, index=False)
    print("✅ Файл успешно сохранён:", output_file)
