import pandas as pd
import re

path:str = 'D:\\4 курс\\Информационные хранилища и аналитические системы\\lab1\\var3\\'
pd.set_option('display.width', None)#отменить усечение столбцов; полный вывод всех столбцов
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)

df_excel3 = pd.read_excel(f"{path}3.xlsx", sheet_name= "3", header=None, names=['Дата', 'Мастер', 'ВИН', 'Услуга', 'Длительность (ч)', 'Цена'], index_col=None, na_values="NA")#заголовок отсутствует - переопределён
df_excel3['Дата'] = pd.to_datetime(df_excel3['Дата']).dt.strftime('%d-%m-%Y')#форматирование даты: '%Y.%m.%d'-> '%d-%m-%Y'
df_excel3['Мастер'] = df_excel3['Мастер'].apply(lambda fio: re.sub(r'(\w)\.(\w)\.$', r'\1. \2.', fio))#Форматирование ФИО: Фамилия И.О. -> Фамилия И. О.
#r - сырая строка; экранирование точки, чтобы найти ее как обычный сивол точки, а не как спец. символ; w - любая буква или цифра(встретился не пробел); () для создания группы и обращения к ней в замене; \1 - первая группа, найденная по шаблону, перед её вставкой вставить пробел; $ - конец строки
#df_excel3['Мастер'] = df_excel3['Мастер'].str.replace(r'\.(\w)', r'. \1', regex=True)#альтернативный способ
df_excel3['Длительность (ч)'] = df_excel3['Длительность (ч)'].replace(to_replace=' ч[.]{0,}', value='', regex=True).astype(int)#убрать приписку длительности в строках и вернуть числовой тип
df_excel3['Цена'] = df_excel3['Цена'].replace(to_replace=' руб[.]{0,}', value='', regex=True).astype(int)#убрать приписку денежных единиц в строках и вернуть числовой тип
df_excel3['Услуга'] = df_excel3['Услуга'].apply(lambda x: x.rstrip())#Удалить пробелы в конце наименований услуг
print(f'excel_data: \n{df_excel3}\n')


df_csv2 = pd.read_csv(f"{path}2.csv", encoding="Windows-1251", header=0, sep=';', index_col=None, na_values="NA")
df_csv2['Паспорт '] = df_csv2['Паспорт '].replace(to_replace='/ ', value='-', regex=True)#форматирование серии и номера: 'серия/ номер' -> 'серия-номер'
df_csv2.columns = df_csv2.columns.str.replace(' ', '')#удалить пробел после Паспорт: 'Паспорт ' -> 'Паспорт '
df_csv2['ФИО'] = df_csv2['ФИО'].apply(lambda fio: ' '.join([fio.split()[0], fio.split()[1][0] + '.', fio.split()[2][0] + '.']))#Форматирование ФИО: Фамилия Имя Отчество -> Фамилия И. О.
df_csv2['Паспорт'] = df_csv2['Паспорт'].apply(lambda x: re.sub(r'(\d{4})\d*-(\d{6})\d*', r'\1-\2', x)) #Форматирование серии и номера паспорта (оставить 4 цифры а серии и 6 цифр в номере)
#df_csv2['Паспорт'] = df_csv2['Паспорт'].apply(lambda x: ''.join(re.findall(r'\d{4}-\d{6}', x)))
print(f'df_csv2: \n{df_csv2}\n')

df_csv4 = pd.read_csv(f"{path}4.csv", encoding="Windows-1251", header=0, sep=';', index_col=None, na_values="NA")
df_csv4['Мастер'] = df_csv4['Мастер'].apply(lambda fio: re.sub(r'(\w)\.(\w)\.$', r'\1. \2.', fio))#форматирование фмо
print(f'df_csv4: \n{df_csv4}\n')


df_txt1 = pd.read_csv(f"{path}1.txt", encoding="UTF-16LE", header=None, names=['Марка','Модель','Год','ВИН'], sep=r'\s+', index_col=None, na_values="NA")#заголовок отсутствует - переопределён; разделитель убирает все пробелы
print(f'df_txt1: \n{df_txt1}\n\n')


df_mark = df_txt1[['Марка']].drop_duplicates()#Создание таблицы марок и удаление дубликатов
df_mark.insert(0, 'id_Марки', df_mark.index + 1)#сгенерировать столбец id
print('\nМарка:\n', df_mark)

df_model = pd.merge(df_txt1[['Марка','Модель']], df_mark, on='Марка', how='inner').drop('Марка', axis='columns')#создание таблицы моделей машин, соединяя таблицу марок с исходным файлом и удалением ненужного столбца
df_model.insert(0, 'id_Модели', df_model.index + 1)
print('\nМодель:\n', df_model)

df_master = df_csv4[['Мастер']].drop_duplicates()#таблица мастеров
df_master.insert(0,'id_Мастера', df_master.index+1)
print('\nМастер:\n', df_master)

df_service = df_excel3[['Услуга']]#Таблица услуг
df_service.insert(0, 'id_Услуги', df_service.index+1)
print('\nУслуга:\n', df_service)

df_client = df_csv2[['Права', 'Паспорт', 'ФИО']].drop_duplicates()#Таблица клиентов
df_client.insert(0, 'id_Клиента', df_client.index+1)
print('\nКлиент:\n', df_client)

df_car = pd.merge(df_txt1, df_model, on='Модель')#таблица машин с id модели и марки
df_car_client_temp = pd.merge(df_txt1, df_csv2, on='ВИН').merge(df_client, on='ФИО').merge(df_car, on='ВИН')#соединить клиентов и их авто по вин, с df_clients для получений id_Клиента
df_car = df_car[['id_Марки', 'id_Модели', 'ВИН', 'Год']]#удалить лишние столбцы
df_car.insert(0, 'id_Авто', df_car.index + 1)
print('\nАвто:\n', df_car)

df_fact = pd.merge(df_excel3, df_csv4, on='Мастер').merge(df_master, on='Мастер').merge(df_car[['id_Авто', 'ВИН', 'Год']], on='ВИН').merge(df_service, on='Услуга').drop(['Мастер', 'Услуга'], axis='columns')#соединение таблиц в факты
print('\nФакт:\n', df_fact)

with pd.ExcelWriter(f"{path}output.xlsx") as writer:
    df_mark.to_excel(writer, index=False, sheet_name='mark')#сохранение в excel
    df_model.to_excel(writer, index=False, sheet_name='model')
    df_client.to_excel(writer, index=False, sheet_name='client')
    df_car.to_excel(writer, index=False, sheet_name='car')
    df_master.to_excel(writer, index=False, sheet_name='master')
    df_service.to_excel(writer, index=False, sheet_name='service')
    df_fact.to_excel(writer, index=False, sheet_name='fact')
