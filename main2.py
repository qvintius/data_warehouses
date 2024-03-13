import pandas as pd
import matplotlib.pyplot as plt

path:str = 'D:\\4 курс\\Информационные хранилища и аналитические системы\\lab\\resources\\lab2\\'
pd.set_option('display.width', None)#отменить усечение столбцов; полный вывод всех столбцов
pd.set_option('display.max_columns', None)

dict_dor_name = pd.read_csv(f"{path}dor_name.csv", encoding="Windows-1251", header=0, sep=',', index_col=None, na_values="NA").set_index('DOR_ID')['NAME'].to_dict()#считывание данных и создание словаря
print(f'dor_name: \n{dict_dor_name}\n')
dict_operiod = pd.read_csv(f"{path}operiod.txt", encoding="Windows-1251", header=0, sep='\t', index_col=None, na_values="NA").set_index('OPERIOD')['NAME'].to_dict()
print(f'operiod: \n{dict_operiod}\n')
df_var = pd.read_sas(f"{path}var.sas7bdat", encoding="Windows-1251")
df_var[['VAR_ID']] = df_var[['VAR_ID']].apply(pd.to_numeric)#преобразование типа данных в числовой
print(f'var: \n{df_var}\n')
dict_var = df_var.set_index('VAR_ID')['NAME'].to_dict()#преобразование в словарь
df_base = pd.read_sas(f"{path}base.sas7bdat", encoding="Windows-1251").query('(VAR_ID=="11410" | VAR_ID=="13740" | VAR_ID=="13760" | VAR_ID=="13060" | VAR_ID=="13070" | VAR_ID=="13090") & (OPERIOD=="H" | OPERIOD=="P")')# фильтрация ненужных данных для ускорения работы
df_base[['DOR_ID', 'VAR_ID']] = df_base[['DOR_ID', 'VAR_ID']].apply(pd.to_numeric)#изменение типа данных на числовой
writer = pd.ExcelWriter(f"{path}output.xlsx", engine='xlsxwriter')  #объект записи в excel
print(f'base: \n{df_base}\n')


report4 = df_base.query('OPERIOD=="H" & VAR_ID==11410').replace({"DOR_ID": dict_dor_name, "VAR_ID": dict_var, "OPERIOD": dict_operiod}).rename(columns={'DOR_ID': 'DOR_NAME', 'VAR_ID': 'VAR_NAME'})#отчет с нужными параметрами и названиями дорог, замена на значения из словаря и переименование столбцов
report4 = report4.groupby([report4['DOR_NAME'], report4['DATE'].map(lambda x: x.year)])['fact'].sum()#суммирование фактов в год, группировка по дорогам и годам
print(report4)
report4 = report4[report4 != 0].unstack().rename_axis('Дорога', axis=0).rename_axis('Год', axis=1)#убирание нулевых значений, разложение, переименование осей
report5 = pd.DataFrame()
report5['Среднее арифметическое'] = report4.transpose().mean()#расчёт среднего арифметического
report5['Среднеквадратичное отклонение'] = report4.transpose().std()#расчет среднеквадратичного отклонения
report4 = report4.cumsum(axis=1)#нарастающая сумма
print(f'report4: \n{report4}\n')
print(f'report5: \n{report5}\n')

report4name = dict_var[11410]#получить значение по которому строится отчет
report4.transpose().plot()
plt.legend(bbox_to_anchor=(1.05, 1))#перемещение легенды
plt.title(report4name)
plt.savefig(f"{path}report4.png", bbox_inches='tight')#запись в png
plt.clf()#сброс графика
report4 = report4.map(lambda x: f'{x:.2f} {'\u20BD'}')#знак рубля unicode
report4.to_excel(writer, sheet_name=report4name[:31])
writer.sheets[report4name[:31]].autofit()#автоширина

report5name='Стат. '+report4name
report5.plot(kind='bar')#столбчатый график
plt.title(report5name)
plt.savefig(f"{path}report5.png", bbox_inches='tight')
plt.clf()
report5.to_excel(writer, sheet_name=report5name[:31])
writer.sheets[report5name[:31]].autofit()#автоширина

report6 = df_base.query('OPERIOD=="H" & (VAR_ID==13740 | VAR_ID==13760)').replace({"DOR_ID": dict_dor_name, "VAR_ID": dict_var, "OPERIOD": dict_operiod}).rename(columns={'DOR_ID': 'DOR_NAME', 'VAR_ID': 'VAR_NAME'})#требуемые данные, замена значениями из словарей
report6 = report6.groupby([report6['DOR_NAME'], report6['DATE'].map(lambda x: x.year), report6['VAR_NAME']])['fact'].sum()
report6 = report6[report6 != 0].unstack(level=-2).rename_axis(('Дорога', 'Значение'), axis=0).rename_axis('Год', axis=1).cumsum(axis=1)#удаление пустых данных, разложение и переименование осей, нарастающий показатель
print(f'report6: \n{report6}\n')
report6name = dict_var[13740] + ' ' + dict_var[13760]#наименование
report6 = report6.map(lambda x: f'{x:.2f} {' ед.'}')
report6.to_excel(writer, sheet_name=report6name[:31])
writer.sheets[report6name[:31]].autofit()

report7 = df_base.query('OPERIOD=="P" & VAR_ID==13060')#требуемые данные
report7 = report7[report7['DATE'].map(lambda x: x.year)==2003]#значения заданного года
report7 = report7[['DATE', 'DOR_ID', 'fact']].rename(columns={'fact': 'fact_a'})#получение нужных столбцов и переименование
report7 = report7[report7 != 0]#получение наименований дорог
report7b = df_base.query('OPERIOD=="P" & VAR_ID==13070')
report7b = report7b[['DATE', 'DOR_ID', 'fact']].rename(columns={'fact': 'fact_b'})
report7c = df_base.query('OPERIOD=="P" & VAR_ID==13090')
report7c = report7c[['DATE', 'DOR_ID', 'fact']].rename(columns={'fact': 'fact_c'})
report7 = report7.merge(report7b, on=['DATE', 'DOR_ID']).merge(report7c, on=['DATE', 'DOR_ID']).replace({"DOR_ID": dict_dor_name}).rename(columns={'DOR_ID': 'DOR_NAME'})#объедиенние показателей и получение значений из словаря
report7 = report7.groupby([report7['DOR_NAME']])[['fact_a', 'fact_b', 'fact_c']].mean().rename_axis('Дорога', axis=0)#группировка по названию и подсчёт среднего
print(f'report7: \n{report7}\n')

report7aname = dict_var[13060]#получение значения названия
report7bname = dict_var[13070]
report7cname = dict_var[13090]
report7 = report7.rename(columns={'fact_a': report7aname, 'fact_b': report7bname, 'fact_c': report7cname})#переимнование столбцов
report7.to_excel(writer, sheet_name=report7aname[:31])
writer.sheets[report7aname[:31]].autofit()
report7.plot(kind='bar')
plt.title(report7aname)
plt.savefig(f"{path}report7.png", bbox_inches='tight')
plt.clf()

writer.close()#запись в файл
