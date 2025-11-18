import pandas as pd
from datetime import datetime, timedelta

EPS = 1e-6  # допуск для сравнения с нулём

###############################################################################
# 1. Парсинг УПД
###############################################################################
def parse_upd(upd_str):
    """
    Парсит строку с УПД формата 'DD.MM.YYYY; НомерУПД; СУММА'.
    Возвращает список словарей:
      [{'date': date, 'doc_number': str, 'amount': float}, ...]
    """
    out = []
    if pd.isna(upd_str):
        return out
    lines = str(upd_str).replace('<br>', '\n').split('\n')
    for line in lines:
        parts = line.split(';')
        if len(parts) < 3:
            continue
        date_str = parts[0].strip()
        doc_num = parts[1].strip()
        amt_str = parts[2].strip()
        try:
            d = datetime.strptime(date_str, '%d.%m.%Y').date()
            a = float(amt_str.replace(' ', '').replace(',', '.').replace('\xa0', ''))
            out.append({
                'date': d,
                'doc_number': doc_num,
                'amount': a
            })
        except:
            continue
    return out


###############################################################################
# 2. Парсинг платежей
###############################################################################
def parse_payments(pay_str):
    """
    Парсит строку с платежами формата 'DD.MM.YYYY; СУММА'.
    Возвращает список словарей:
      [{'date': date, 'amount': float}, ...]
    """
    out = []
    if pd.isna(pay_str):
        return out
    lines = str(pay_str).replace('<br>', '\n').split('\n')
    for line in lines:
        if ';' not in line:
            continue
        date_str, amt_str = line.split(';', 1)
        try:
            d = datetime.strptime(date_str.strip(), '%d.%m.%Y').date()
            a = float(amt_str.strip().replace(' ', '').replace(',', '.').replace('\xa0', ''))
            out.append({
                'date': d,
                'amount': a
            })
        except:
            continue
    return out


###############################################################################
# 3. Распределение оплат по УПД (FIFO)
###############################################################################
def distribute_payments(upds, payments):
    """
    Распределяет оплаты по УПД в порядке их даты (FIFO).
    Возвращает словарь: {doc_number: [{'payment_date': date, 'amount': float}, ...]}

    Логика:
      1) Сортируем УПД по возрастанию даты.
      2) Сортируем платежи по возрастанию даты.
      3) Идём по каждому платежу, "закрывая" УПД по очереди, пока не израсходуем
         весь платеж или не закончатся УПД.
    """
    upds_sorted = sorted(upds, key=lambda x: x['date'])
    payments_sorted = sorted(payments, key=lambda x: x['date'])

    distribution = {upd['doc_number']: [] for upd in upds_sorted}
    remaining_upds = {upd['doc_number']: upd['amount'] for upd in upds_sorted}

    for pay in payments_sorted:
        pay_amount = pay['amount']
        for upd in upds_sorted:
            doc_num = upd['doc_number']
            if remaining_upds[doc_num] > 0 and pay_amount > 0:
                allocated = min(pay_amount, remaining_upds[doc_num])
                distribution[doc_num].append({
                    'payment_date': pay['date'],
                    'amount': allocated
                })
                remaining_upds[doc_num] -= allocated
                pay_amount -= allocated
                if pay_amount <= 0:
                    break

    return distribution


###############################################################################
# 4. Расчет коммерческого кредита (КК) для одного УПД
###############################################################################
def compute_cc_for_upd(upd_info, payments_for_upd, daily_rate, current_date):
    """
    Учитываем предоплату (платежи до или в день отгрузки):
      1) Суммируем все платежи, у которых payment_date <= upd_date (предоплата),
         и уменьшаем на них базу УПД (без начисления процентов).
      2) Оставшуюся часть начинаем "процентовать" с upd_date+1.
      3) Идём по платежам, которые приходят после даты отгрузки, и начисляем
         проценты за период между предыдущим "срезом" и датой оплаты.
      4) Если остаток > 0 в конце, начисляем проценты до current_date.

    Параметры:
      - upd_info = {'date': date, 'amount': float, 'doc_number': str}
      - payments_for_upd = [{'payment_date': date, 'amount': float}, ...]
      - daily_rate = дневная ставка (напр. 0.003 для 0.3%)
      - current_date = текущая дата (конец периода)

    Возвращает кортеж: (общая сумма процентов, список деталей по периодам).
    """
    upd_date = upd_info['date']
    upd_amount = upd_info['amount']
    doc_number = upd_info['doc_number']

    # Сортируем платежи, чтобы шли по возрастанию даты
    payments_for_upd = sorted(payments_for_upd, key=lambda x: x['payment_date'])

    remaining = upd_amount
    total_interest = 0.0
    details = []

    # 1) Предоплата: все платежи, дата которых <= upd_date
    for pay in payments_for_upd:
        if pay['payment_date'] <= upd_date:
            allocated = min(remaining, pay['amount'])
            remaining -= allocated
            # Если уже весь УПД погашен предоплатой — проценты = 0
            if remaining <= EPS:
                remaining = 0.0
                break

    if abs(remaining) < EPS:
        # Полностью погашено до или в день отгрузки, процентов нет
        return (0.0, [{
            'doc_number': doc_number,
            'base': 0.0,
            'start_date': upd_date,
            'end_date': upd_date,
            'days': 0,
            'interest': 0.0
        }])

    # 2) Начинаем считать проценты с дня, следующего за отгрузкой
    base_date = upd_date
    accrual_start = base_date + timedelta(days=1)
    if accrual_start > current_date:
        # Если текущая дата раньше, чем accrual_start, то проценты не начисляются
        return (0.0, [{
            'doc_number': doc_number,
            'base': round(remaining, 2),
            'start_date': accrual_start,
            'end_date': accrual_start,
            'days': 0,
            'interest': 0.0
        }])

    # 3) Обработка платежей, пришедших ПОСЛЕ даты отгрузки
    for pay in payments_for_upd:
        pay_date = pay['payment_date']
        if pay_date <= upd_date:
            # Эти платежи уже учли как предоплату
            continue
        if pay_date <= base_date:
            # Если дата платежа не продвигает базовую дату, пропускаем
            continue

        # Если остаток уже 0, нет смысла продолжать
        if remaining <= EPS:
            break

        # Кол-во дней: (pay_date - (base_date+1)) + 1
        days_count = (pay_date - (base_date + timedelta(days=1))).days + 1
        if days_count > 0:
            interest = remaining * daily_rate * days_count
            total_interest += interest
            details.append({
                'doc_number': doc_number,
                'base': round(remaining, 2),
                'start_date': base_date + timedelta(days=1),
                'end_date': pay_date,
                'days': days_count,
                'interest': round(interest, 2)
            })

        # Уменьшаем остаток на сумму платежа
        allocated = min(remaining, pay['amount'])
        remaining -= allocated
        # Сдвигаем базовую дату на день платежа
        base_date = pay_date

    # 4) Если после всех оплат остаток > 0, считаем проценты до current_date
    if remaining > EPS and base_date < current_date:
        days_count = (current_date - (base_date + timedelta(days=1))).days + 1
        if days_count > 0:
            interest = remaining * daily_rate * days_count
            total_interest += interest
            details.append({
                'doc_number': doc_number,
                'base': round(remaining, 2),
                'start_date': base_date + timedelta(days=1),
                'end_date': current_date,
                'days': days_count,
                'interest': round(interest, 2)
            })

    return (round(total_interest, 2), details)


###############################################################################
# 5. Основной скрипт
###############################################################################
def main():
    # Пути к файлам (пример)
    input_path = r"C:\Users\nkazakov\Downloads\тест для Питона.xlsx"
    output_path = r"C:\Users\nkazakov\Downloads\Обработанный_итоговый_файл.xlsx"

    # Читаем данные
    df = pd.read_excel(input_path, sheet_name='TDSheet', dtype={'ИНН': str})

    all_debug_records = []

    for idx, row in df.iterrows():
        # (A) Парсим данные
        upd_list = parse_upd(row.get('УПД'))
        fact_list = parse_payments(row.get('Фактические оплаты по датам'))

        # (B) Определяем ставку коммерческого кредита
        raw_rate = row.get('Процент коммерческого кредита')
        if pd.isna(raw_rate):
            daily_rate = 0.003  # 0.3% в день по умолчанию
        else:
            daily_rate = float(raw_rate) / 100.0

        # (C) Распределяем оплаты по УПД (FIFO)
        distribution = distribute_payments(upd_list, fact_list)

        # Текущая дата (дата запуска скрипта)
        current_date = datetime.today().date()

        total_cc = 0.0

        # (D) Считаем коммерческий кредит для каждого УПД
        for upd_info in upd_list:
            payments_for_upd = distribution.get(upd_info['doc_number'], [])
            cc_val, det_list = compute_cc_for_upd(
                upd_info,
                payments_for_upd,
                daily_rate,
                current_date
            )
            total_cc += cc_val

            # Добавляем детализацию с дополнительной информацией
            for seg in det_list:
                seg['row_index'] = idx
                seg['Контрагент'] = row.get('Контрагент', '')
                seg['Договор'] = row.get('Договор', '')
            all_debug_records.extend(det_list)

        # (E) Если сумма оплат >= сумме отгрузки – обнуляем КК (логика "если всё оплачено")
        total_shipped = sum(upd['amount'] for upd in upd_list)
        total_paid = sum(p['amount'] for p in fact_list)
        if total_paid >= total_shipped:
            total_cc = 0.0
            # Удаляем старую детализацию
            all_debug_records = [rec for rec in all_debug_records if rec['row_index'] != idx]
            # Добавляем запись, что процентов нет
            for upd in upd_list:
                all_debug_records.append({
                    'row_index': idx,
                    'doc_number': upd['doc_number'],
                    'base': 0,
                    'start_date': upd['date'],
                    'end_date': upd['date'],
                    'days': 0,
                    'interest': 0.0,
                    'Контрагент': row.get('Контрагент', ''),
                    'Договор': row.get('Договор', '')
                })

        # Записываем итоговый коммерческий кредит в датафрейм
        df.at[idx, 'Коммерческий кредит'] = round(total_cc, 2)

    # (F) Формируем таблицу детализации
    df_debug = pd.DataFrame(all_debug_records)
    df_debug.sort_values(by=['row_index', 'doc_number', 'start_date'], inplace=True)

    # (G) Сохраняем результат
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Сводный отчет', index=False)
        df_debug.to_excel(writer, sheet_name='Детализация', index=False)

    print("Файл сохранён:", output_path)


if __name__ == '__main__':
    main()
