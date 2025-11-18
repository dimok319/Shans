import pandas as pd
import re
from datetime import datetime, date, timedelta

# Очистка суммы
def clean_amount_string(amt_str):
    amt_str = str(amt_str).replace("\u00A0", "").replace("\u2007", "").replace("\u202F", "")
    amt_str = re.sub(r"[^0-9\.,-]+", "", amt_str)
    return amt_str.replace(",", ".")

# Парсинг оплаты
def parse_payment_info(cell, mode="plan", filter_past=False, today=None):
    if pd.isna(cell) or str(cell).strip() == "":
        return 0, None
    lines = str(cell).splitlines()
    records = []
    for line in lines:
        parts = re.split(r"\s*;\s*", line)
        if len(parts) < 2:
            continue
        date_str = parts[0].strip()
        amt_str = clean_amount_string(parts[1].strip())
        try:
            dt = datetime.strptime(date_str, "%d.%m.%Y").date()
            amt = float(amt_str)
        except:
            continue
        records.append((dt, amt))
    if mode == "plan" and filter_past:
        valid = [(dt, amt) for dt, amt in records if dt and dt <= today]
        total_amount = sum(amt for dt, amt in valid)
        date_val = min([dt for dt, amt in valid], default=None)
    else:
        total_amount = sum(amt for dt, amt in records)
        if mode == "plan":
            date_val = min([dt for dt, amt in records if dt], default=None)
        else:
            date_val = max([dt for dt, amt in records if dt], default=None)
    return total_amount, date_val

# Поиск колонки
def find_column_by_keywords(df, keywords):
    keywords = [k.lower().strip() for k in keywords]
    matches = [col for col in df.columns if all(k in col.lower() for k in keywords)]
    if not matches:
        raise KeyError(f"Столбец с ключами {keywords} не найден.")
    return matches[0]

# Дни просрочки
def calculate_group_overdue_days(row):
    try:
        plan_text = row.get("Оплаты по дням (план)")
        fact_text = row.get("Оплата по дням (факт)")
        if pd.isna(plan_text) or not isinstance(plan_text, str):
            return 0

        # Плановые оплаты
        plan_records = []
        for line in plan_text.splitlines():
            parts = re.split(r"\s*;\s*", line)
            if len(parts) >= 2:
                try:
                    dt = datetime.strptime(parts[0].strip(), "%d.%m.%Y").date()
                    amt = float(clean_amount_string(parts[1].strip()))
                    plan_records.append([dt, amt])
                except:
                    continue

        # Фактические оплаты
        fact_records = []
        if isinstance(fact_text, str):
            for line in fact_text.splitlines():
                parts = re.split(r"\s*;\s*", line)
                if len(parts) >= 2:
                    try:
                        dt = datetime.strptime(parts[0].strip(), "%d.%m.%Y").date()
                        amt = float(clean_amount_string(parts[1].strip()))
                        fact_records.append([dt, amt])
                    except:
                        continue

        plan_records.sort()
        fact_records.sort()
        total_days = 0

        for plan_date, plan_amt in plan_records:
            if plan_amt <= 0:
                continue

            remaining_amt = plan_amt
            i = 0

            while i < len(fact_records) and remaining_amt > 0:
                pay_date, pay_amt = fact_records[i]
                apply_amt = min(remaining_amt, pay_amt)

                if pay_date <= plan_date:
                    remaining_amt -= apply_amt  # вовремя — без просрочки
                else:
                    days = (pay_date - plan_date).days
                    total_days += days  # оплачено с просрочкой
                    remaining_amt -= apply_amt

                fact_records[i][1] -= apply_amt
                if fact_records[i][1] <= 0:
                    i += 1

            # если часть не оплачена вообще
            if remaining_amt > 0 and plan_date < today:
                days = (today - plan_date).days
                total_days += days

        return total_days
    except:
        return 0

# Проценты по каждой просроченной дате (группировка)
def calculate_group_percentage(row):
    try:
        plan_text = row.get("Оплаты по дням (план)")
        fact_text = row.get("Оплата по дням (факт)")
        if pd.isna(plan_text) or not isinstance(plan_text, str):
            return 0

        # Плановые оплаты
        plan_records = []
        for line in plan_text.splitlines():
            parts = re.split(r"\s*;\s*", line)
            if len(parts) >= 2:
                try:
                    dt = datetime.strptime(parts[0].strip(), "%d.%m.%Y").date()
                    amt = float(clean_amount_string(parts[1].strip()))
                    plan_records.append([dt, amt])
                except:
                    continue

        # Фактические оплаты
        fact_records = []
        if isinstance(fact_text, str):
            for line in fact_text.splitlines():
                parts = re.split(r"\s*;\s*", line)
                if len(parts) >= 2:
                    try:
                        dt = datetime.strptime(parts[0].strip(), "%d.%m.%Y").date()
                        amt = float(clean_amount_string(parts[1].strip()))
                        fact_records.append([dt, amt])
                    except:
                        continue

        plan_records.sort()
        fact_records.sort()
        total_interest = 0

        for plan_date, plan_amt in plan_records:
            if plan_amt <= 0:
                continue

            remaining_amt = plan_amt
            i = 0

            while i < len(fact_records) and remaining_amt > 0:
                pay_date, pay_amt = fact_records[i]
                apply_amt = min(remaining_amt, pay_amt)

                if pay_date <= plan_date:
                    # Оплачено в срок — проценты не начисляем
                    remaining_amt -= apply_amt
                else:
                    # Оплачено с просрочкой
                    delay_days = (pay_date - plan_date).days
                    interest = apply_amt * (0.28 / 365) * delay_days
                    total_interest += interest
                    remaining_amt -= apply_amt

                fact_records[i][1] -= apply_amt
                if fact_records[i][1] <= 0:
                    i += 1

            # Если осталась непогашенная сумма — считаем проценты до today
            if remaining_amt > 0 and plan_date < today:
                delay_days = (today - plan_date).days
                interest = remaining_amt * (0.28 / 365) * delay_days
                total_interest += interest

        return round(total_interest, 2)
    except:
        return 0

# Статус оплаты
def aggregated_status(row):
    try:
        plan_date = row.get("PlanDatePast")
        fact_date = row.get("FactDateMax")
        plan_amt = row.get("PlanAmountPast", 0)
        fact_amt = row.get("FactAmountTotal", 0)
        overdue_days = row.get("Агрегированные дни просрочки", 0)
        full_date = row.get("PlanDateFull")

        if pd.isna(plan_date):
            return "Плановая дата не наступила"

        plan_date = pd.to_datetime(plan_date).date()

        # Если дата плана = сегодня — ещё можно оплатить
        if plan_date == today:
            return "Плановая дата не наступила"

        # Если дата плана > сегодня
        if plan_date > today:
            if fact_amt >= row.get("PlanAmountFull", 0):
                return "Все оплачено в срок"
            return "Плановая дата не наступила"

        # Если факт = 0 и дата прошла — это долг
        if fact_amt == 0 and plan_date < today:
            return "Есть долг"

        # Всё оплачено, без просрочек
        if row["Агрегированный долг"] == 0 and overdue_days == 0:
            return "Все оплачено в срок"

        # Всё оплачено, но с просрочкой
        if row["Агрегированный долг"] == 0 and overdue_days > 0:
            return "Все оплачено, но с просрочкой"

        # Частично оплачено, но дата прошла
        if row["Агрегированный долг"] > 0 and full_date > today:
            return "Не все оплачено, просрочка"

        return "Есть долг"
    except:
        return "Ошибка в расчёте статуса"

# === MAIN ===
if __name__ == "__main__":
    today = datetime.today().date()

    input_file = r"\\192.168.1.211\дебиторская задолженность\Отчеты\Выгрузка 1С\ДЗ_1С_НОВЫЙ (XLSX).xlsx"
    output_file = r"C:\Users\nkazakov\Downloads\Должник по ДЗ.xlsx"

    df = pd.read_excel(input_file, header=2)
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.duplicated()]

    df["Заказ"] = df["Заказ клиента"].astype(str).str.strip()
    col_plan_name = find_column_by_keywords(df, ["по дням", "план"])
    col_fact_name = find_column_by_keywords(df, ["по дням", "факт"])

    df[["PlanAmountPast", "PlanDatePast"]] = df[col_plan_name].apply(
        lambda x: pd.Series(parse_payment_info(x, "plan", True, today)))
    df[["PlanAmountFull", "PlanDateFull"]] = df[col_plan_name].apply(
        lambda x: pd.Series(parse_payment_info(x, "plan", False)))
    df[["FactAmountTotal", "FactDateMax"]] = df[col_fact_name].apply(
        lambda x: pd.Series(parse_payment_info(x, "fact")))

    df["Долг"] = (df["PlanAmountPast"] - df["FactAmountTotal"]).clip(lower=0)
    df["Оплата по днями (факт).руб."] = df["FactAmountTotal"]
    df["Проценты"] = df.apply(calculate_group_percentage, axis=1)

    # Привязка регионов к дивизионам
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

    grouped = df.groupby("Заказ", as_index=False).agg({
        "PlanAmountPast": "sum", "FactAmountTotal": "sum", "PlanDatePast": "min", "FactDateMax": "max",
        "PlanDateFull": "max", "Дивизион": "first", "Регион.Наименование": "first", "Сезон": "first",
        "Контрагент.Сокращенное юр. наименование": "first", "Контрагент.ИНН": "first",
        "Канал продаж": "first", "Вид продаж": "first", "Сумма оплаты по заказу, (руб.)": "first",
        "Оплата по днями (факт).руб.": "first"
    })

    grouped["Агрегированный долг"] = (grouped["PlanAmountPast"] - grouped["FactAmountTotal"]).clip(lower=0)
    grouped = grouped.merge(df[["Заказ", col_plan_name, col_fact_name]], on="Заказ", how="left")
    grouped["Агрегированные дни просрочки"] = grouped.apply(calculate_group_overdue_days, axis=1)
    grouped["Агрегированные проценты"] = grouped.apply(calculate_group_percentage, axis=1)
    grouped["Агрегированный статус оплаты"] = grouped.apply(aggregated_status, axis=1)

    final_columns = [
        "Дивизион", "Регион.Наименование", "Сезон", "Контрагент.Сокращенное юр. наименование", "Контрагент.ИНН",
        "Канал продаж", "Вид продаж", "Сумма оплаты по заказу, (руб.)", "Оплата по днями (факт).руб.",
        "PlanAmountPast", "FactAmountTotal", "PlanDatePast", "FactDateMax", "PlanDateFull",
        "Агрегированный долг", "Агрегированные дни просрочки", "Агрегированные проценты",
        "Агрегированный статус оплаты", "Заказ", col_plan_name, col_fact_name
    ]

    grouped = grouped[final_columns]
    grouped.to_excel(output_file, index=False)
    print("✅ Файл успешно сохранён.")
