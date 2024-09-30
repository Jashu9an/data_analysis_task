import pandas as pd
import matplotlib.pyplot as plt

# Загрузка данных из Excel файла
data = pd.read_excel('data.xlsx')

# Преобразование столбца 'receiving_date' в формат даты
data['receiving_date'] = pd.to_datetime(data['receiving_date'], errors='coerce')

# Фильтрация данных по месяцам
May_data        = data[data[data['status'] == 'Май 2021'].index[0] : data[data['status'] == 'Июнь 2021'].index[0]]
June_data       = data[data[data['status'] == 'Июнь 2021'].index[0] : data[data['status'] == 'Июль 2021'].index[0]]
July_data       = data[data[data['status'] == 'Июль 2021'].index[0] : data[data['status'] == 'Август 2021'].index[0]]
August_data     = data[data[data['status'] == 'Август 2021'].index[0] : data[data['status'] == 'Сентябрь 2021'].index[0]]
September_data  = data[data[data['status'] == 'Сентябрь 2021'].index[0] : data[data['status'] == 'Октябрь 2021'].index[0]]
October_data    = data[data[data['status'] == 'Октябрь 2021'].index[0]:]

# 1. Общая выручка за июль 2021 по непросроченным сделкам
total_revenue_july = 0
for index, row in July_data.iterrows():
    if (row['status'] != 'ПРОСРОЧЕНО') & (row['status'] != 'Июль 2021'):
        total_revenue_july += float(row['sum'])
print(f'Общая выручка за июль 2021: {total_revenue_july:.2f}')

# 2. Изменение выручки компании за рассматриваемый период
revenue_may = 0
for index, row in May_data.iterrows():
    if (row['status'] != 'Май 2021'):
        revenue_may += float(row['sum'])

revenue_june = 0
for index, row in June_data.iterrows():
    if (row['status'] != 'Июнь 2021'):
        revenue_june += float(row['sum'])

revenue_july = 0
for index, row in July_data.iterrows():
    if (row['status'] != 'Июль 2021'):
        revenue_july += float(row['sum'])

revenue_august= 0
for index, row in August_data.iterrows():
    if (row['status'] != 'Август 2021'):
        revenue_august += float(row['sum'])

revenue_september = 0
for index, row in September_data.iterrows():
    if (row['status'] != 'Сентябрь 2021'):
        revenue_september += float(row['sum'])

revenue_october = 0
for index, row in October_data.iterrows():
    if (row['status'] != 'Октябрь 2021'):
        revenue_october += float(row['sum'])

def addlabels(x,y):
    for i in range(len(x)):
        plt.text(i, y[i], round(y[i], 2), ha = 'center')

months = ['Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь']
plt.bar(months, [revenue_may, revenue_june, revenue_july, revenue_august, revenue_september, revenue_october])
addlabels(months, [revenue_may, revenue_june, revenue_july, revenue_august, revenue_september, revenue_october])
plt.title('Выручка компании по месяцам')
plt.xlabel('Месяц')
plt.ylabel('Выручка')

# 3. Менеджеры, привлечшие больше всего средств в сентябре 2021
manager_revenue_september = September_data.groupby('sale')['sum'].sum()
top_manager_september = manager_revenue_september.idxmax()
top_revenue_september = manager_revenue_september.max()
print(f'Менеджер, привлечший больше всего средств в сентябре 2021: {top_manager_september} с выручкой {top_revenue_september}')

# 4. Преобладающий тип сделок в октябре 2021
deal_type_counts = October_data['new/current'].value_counts()
dominant_deal_type_october = deal_type_counts.idxmax()
print(f'Преобладающий тип сделок в октябре 2021: {dominant_deal_type_october}')

# 5. Количество оригиналов договора по майским сделкам, полученным в июне 2021
originals_received_june = 0
for index, row in May_data.iterrows():
    if (row['receiving_date'] >= pd.to_datetime('01-06-2021', format='%d-%m-%Y')) & (row['receiving_date'] <= pd.to_datetime('30-06-2021', format='%d-%m-%Y')):
        originals_received_june += 1
print(f'Количество оригиналов договора по майским сделкам, полученным в июне 2021: {originals_received_june}')

# Расчет бонусов для менеджеров на 01.07.2021
july_2021_bonuses = {}

for index, row in data[: data[data['status'] == 'Июль 2021'].index[0]].iterrows():
    if (row['sale'] == '') or (row['sale'] == '-'):
        continue
    else:
        manager = row['sale']
        if row['new/current'] == 'новая':
            if row['status'] == 'ОПЛАЧЕНО' and row['document'] == 'оригинал':
                if (row['receiving_date'] >= pd.to_datetime('01-05-21', format = '%d-%m-%y')) & (row['receiving_date'] <= pd.to_datetime('30-06-21', format = '%d-%m-%y')):
                    bonus = row['sum'] * 0.07
                else:
                    bonus = 0
        elif row['new/current'] == 'текущая':
            if row['status'] != 'ПРОСРОЧЕНО' and row['document'] == 'оригинал':
                if (row['receiving_date'] >= pd.to_datetime('01-05-21', format = '%d-%m-%y')) & (row['receiving_date'] <= pd.to_datetime('30-06-21', format = '%d-%m-%y')):
                    if row['sum'] > 10000:
                        bonus = row['sum'] * 0.05
                    else:
                        bonus = row['sum'] * 0.03
            else:
                bonus = 0
        else:
            bonus = 0

        if manager in july_2021_bonuses:
            july_2021_bonuses[manager] += bonus
        else:
            july_2021_bonuses[manager] = bonus

print('Бонусы менеджеров на 01.07.2021:')
for manager, bonus in july_2021_bonuses.items():
    print(f'{manager}: {bonus:.2f}')

plt.show()