import pandas as pd

# Загрузка данных из файлов и предварительное преобразование
def load_and_prepare_data(file_path):
    data = pd.read_excel(file_path)
    data['Дата'] = pd.to_datetime(data['Дата']).dt.strftime('%Y-%m-%d')
    data['Комиссия'] = data['Комиссия'].abs()
    data.sort_values(by='RRN', inplace=True)
    return data

ioka_report = load_and_prepare_data('iokas_report.xlsx')
banks_report = load_and_prepare_data('banks_report.xlsx')

# Объединение данных по RRN
merged_data = pd.merge(ioka_report, banks_report, on='RRN', suffixes=('_ioka', '_bank'), how='outer')

# Нахождение столбцов с расхождениями
mismatched_columns = [col for col in ['Дата', 'Сумма', 'Комиссия'] if not merged_data[f'{col}_ioka'].equals(merged_data[f'{col}_bank'])]

# Функция для определения причины расхождения
def get_mismatch_reason(row):
    reasons = []
    
    if pd.isna(row['Дата_bank']):
        reasons.append('Платеж отсутствует в отчете Банка')
    if row['Дата_ioka'] != row['Дата_bank']:
        reasons.append('Расхождение в столбце "Дата"')
    if row['Сумма_ioka'] != row['Сумма_bank']:
        reasons.append('Расхождение в столбце "Сумма"')
    if row['Комиссия_ioka'] != row['Комиссия_bank']:
        reasons.append('Расхождение в столбце "Комиссия"')

    if reasons:
        return ', '.join(reasons)
    else:
        return 'Неизвестная причина'

# Добавление столбца с причинами расхождения
merged_data['Причина расхождения'] = merged_data.apply(get_mismatch_reason, axis=1)

# Фильтрация расхождений и совпадающих платежей
mismatched_payments = merged_data[(merged_data['RRN'].isnull()) | (merged_data['Причина расхождения'] != 'Неизвестная причина')]
matched_payments = merged_data[(merged_data['Причина расхождения'] == 'Неизвестная причина')]

# Создание нового Excel файла
with pd.ExcelWriter('result.xlsx', engine='openpyxl') as writer:
    matched_payments.drop(columns=['Причина расхождения']).to_excel(writer, sheet_name='Совпадающие платежи', index=False)
    mismatched_payments.to_excel(writer, sheet_name='Расхождения и причины', index=False)

print("Готово! Результаты сохранены в файле result.xlsx")
