#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import seaborn as sns


# In[2]:


# загружаем данные
df_sale = pd.read_excel('C:/Users/olap_cubes.xlsx',sheet_name='Лист1')
df_stores = pd.read_excel('C:/Users/sql.xlsx', index_col=0, usecols='A,B,D,K',names=['Code','OpenDate','CloseDate','Region'])


# In[3]:


# форматируем, фильтруем и группируем данные с учетом необходимых для анализа дат
def last_week(df_sale, d1, d2):
    df_sale['Date'] = pd.to_datetime(df_sale['Date'],format="%d.%m.%Y")
    df_sale1 = df_sale[(df_sale['Date']>=d1) & (df_sale['Date']<=d2)]
    return df_sale1.groupby('Store').sum()

# форматируем, фильтруем и группируем данные для таблицы продаж предыдущих лет 
def lfl(df_sale, d1, d2):
    df_sale['Date'] = pd.to_datetime(df_sale['Date'],format="%d.%m.%Y")
    df_sale['curDate'] = df_sale['Date'] + timedelta(days=728) # нахождение аналогичной даты
    last_sale1 = df_sale[(df_sale['curDate']>=d1) & (df_sale['curDate']<=d2)][['Store','Sales','Date','curDate']]
    last_sale2 = pd.merge(last_sale1,df_sale[['Store','Date','Sales']], left_on=['Store','curDate'], right_on=['Store','Date']) #объединяем продажи за текущий и прошлый аналогичный период
    last_sale3 = last_sale2[['Store','Sales_x','Sales_y','Date_y']].merge(df_stores, how='inner', left_on='Store', right_index=True)
    last_sale = last_sale3[(last_sale3['Sales_x']>0) & (last_sale3['Sales_y']>0) & (last_sale3['Date_y']-last_sale3['OpenDate']>timedelta(days=786))].groupby('Region').sum()
    last_sale['LFL']=(last_sale['Sales_y']/last_sale['Sales_x']-1).apply(lambda x:'{:.0%}'.format(x))
    return last_sale[['Sales_y','Sales_x','LFL']]
    
# форматируем, фильтруем и группируем данные для таблицы продаж предыдущей недели 
def prelast_week(df_sale, d1, d2):
    df_sale['Date'] = pd.to_datetime(df_sale['Date'],format="%d.%m.%Y")
    df_sale['curDate'] = df_sale['Date'] + timedelta(days=7) # нахождение аналогичной даты
    last_sale1 = df_sale[(df_sale['curDate']>=d1) & (df_sale['curDate']<=d2)][['Store','Sales','Date','curDate']]
    last_sale2 = pd.merge(last_sale1,df_sale[['Store','Date','Sales']], left_on=['Store','curDate'], right_on=['Store','Date']) #объединяем продажи за текущий и прошлый аналогичный период
    last_sale3 = last_sale2[['Store','Sales_x','Sales_y','Date_y']].merge(df_stores, how='inner', left_on='Store', right_index=True)
    last_sale = last_sale3[(last_sale3['Sales_x']>0) & (last_sale3['Sales_y']>0)].groupby('Region').sum()
    last_sale['%LW']=(last_sale['Sales_y']/last_sale['Sales_x']-1).apply(lambda x:'{:.0%}'.format(x))
    last_sale = last_sale.rename(columns={'Sales_y':'Sales_CW','Sales_x':'Sales_LW'})
    return last_sale[['Sales_CW','Sales_LW','%LW']]

# таблица дат для графика
def chart(df_sale, d1, d2):
    calr = {1:31,2:28,3:31,4:30,5:31,6:30,7:31,8:31,9:30,10:31,11:30,12:31}
    df_sale['Date'] = pd.to_datetime(df_sale['Date'],format="%d.%m.%Y")
    df_sale1 = df_sale[(df_sale['Date']>=d1.replace(day=1)) & (df_sale['Date']<=d2.replace(day=calr[d2.month]))]
    return df_sale1.groupby('Date').sum()


# In[4]:


week_begin = datetime(2023,6,1) #определяем даты для анализа

week_end = week_begin + timedelta(days=6)

sales = last_week(df_sale,week_begin,week_end) #вызываем функцию форматирования, фильтрации и группировки таблицы продаж по датам
sales1 = sales.merge(df_stores, how='inner', left_index=True, right_index=True) #объединение таблиц
sales1 = sales1[sales1['Sales']>0] #очищаем данные без продаж
sales2 = sales1.groupby('Region').sum() #группируем по регионам
sales2_1 = sales1.groupby('Region').count() #группируем по регионам
last_sale = lfl(df_sale,week_begin,week_end) #вызываем функцию добавления таблицы с данными предыдущих дат
prelastweek = prelast_week(df_sale,week_begin,week_end) #вызываем функцию добавления таблицы с данными предыдущей недели


# In[5]:


sales2['%Sale'] = (sales2['Sales']/sum(sales2['Sales'])).apply(lambda x:'{:.0%}'.format(x)) #рассчет доли продаж в руб.
sales2['Units'] = list(map(int,sales2['Units'])) #рассчет продажи шт.
sales2 = sales2.merge(sales2_1['Year'], how='inner', left_index=True, right_index=True) #добавление столбца кол-ва магазинов
sales2['Qty stores'] = sales2['Year_y']
sales2 = sales2.merge(last_sale, how='inner', left_index=True, right_index=True) #добавление столбца lfl
sales2 = sales2.merge(prelastweek, how='inner', left_index=True, right_index=True) #добавление столбца %LW
sales2.loc['TOTAL'] = sales2.agg({'Qty stores':'sum','Units':'sum','Trans Netto':'sum','Sales':'sum','Plans':'sum'}) #добавление строки итогов
sales2.at['TOTAL','LFL'] = '{:.0%}'.format(sales2['Sales_y'].sum()/sales2['Sales_x'].sum()-1) #рассчет итога lfl
sales2.at[:,'%Plan Impl'] = (sales2['Sales']/sales2['Plans']).apply(lambda x:'{:.0%}'.format(x)) #рассчет выполнение плана
sales2.loc[:,'ATV'] = round(sales2['Sales']/sales2['Trans Netto'],1) #рассчет среднего чека в руб.
sales2.loc[:,'UPT'] = round(sales2['Units']/sales2['Trans Netto'],2) #рассчет среднего чека в шт.
sales2.at['TOTAL','%LW'] = '{:.0%}'.format(sales2['Sales_CW'].sum()/sales2['Sales_LW'].sum()-1) #рассчет итога %lw

sales_itog = sales2[['Qty stores','%Sale','%LW','LFL','Units','%Plan Impl','ATV','UPT']] #формирование итоговой таблицы
sales_itog = sales_itog.fillna('')


# In[6]:


#форматирование таблицы, добавление условного форматирования
sales_itog['Qty stores'] = list(map(int,sales_itog['Qty stores']))
sales_itog['Units'] = list(map(int,sales_itog['Units']))
headers1 = {
    'selector': 'thead',
    'props': 'background-color: #C65911; color: white;font-weight:normal;text-align: center;'}
headers2 = {
    'selector': 'th.row_heading',
    'props': 'background-color: white; color: black;font-weight:normal;text-align: center;'}
cell = {
    'selector': '.row3',
    'props': 'background-color: #FCE4D6; font-weight:bold;'}
sales_itog = sales_itog.style.set_table_styles([headers1,headers2,cell]).bar(color=['#C6E0B4','#C6E0B4'], subset=pd.IndexSlice['Region1':'Region3','Units']).background_gradient(cmap='YlGn', subset=pd.IndexSlice['Region1':'Region3','ATV']).format(precision=1, thousands=" ", decimal=",")


# In[12]:


#создаем график для отчета
plt.style.use('seaborn-whitegrid')
fig, ax = plt.subplots(figsize=(14, 3)) 
chart_table = chart(df_sale,week_begin,week_end) #вызываем таблицу дат для графика
dates = chart_table.index 
dates = [str(int(str(i)[-11:-9]))+'.'+str(i)[-13:-12] for i in dates]
sales = chart_table ['Sales']
plans = chart_table ['Plans']
ax.plot(dates,sales, color='#ED7D31',marker='o',label='Sales') #добавляем данные в график
ax.plot(dates,plans, color='#5B9BD5',marker='o',label='Plans')

month_en = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'May',6:'June',7:'July',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
if week_begin.month==week_end.month: #добавляем название графика
    ax.set_title('Dynamics Plan&Fact Sales for '+str(month_en[week_begin.month]),fontsize=18) 
else:    
    ax.set_title('Dynamics Plan&Fact Sales for '+str(month_en[week_begin.month])+'-'+str(month_en[week_end.month]),fontsize=18) 
    
ax.legend() #добавляем легенду
#ax.grid() #удаляем сетку
ax.set_yticks([]) #удаляем осьУ 


plt.show()
print(f'\033[1mSales period = {week_begin.strftime("%d %B")} - {week_end.strftime("%d %B")}\033[0m')
sales_itog

