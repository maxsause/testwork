import datetime

import pandas as pd  # pip install pandas openpyxl
import plotly.express as px  # pip install plotly-express
import streamlit as st  # pip install streamlit


st.set_page_config(
    page_title="",
    layout="wide"
)
st.title('Дашборд по тестовому заданию')
st.divider()


# 1
st.subheader('Общее кол-во сотрудников из выбранного подразделения, за выбранный период:')

col1, col2, col3 = st.columns([1, 1, 3])
with col1:
    division = st.selectbox(
        'Выберите подразделение',
        ('Подразделение1', 'Подразделение2')
    )
df_all1 = pd.ExcelFile('Подразделение1.xlsx')
df_all2 = pd.ExcelFile('Подразделение2.xlsx')

today = datetime.datetime.now()
jan_1 = datetime.date(today.year - 1, 1, 1)
dec_31 = datetime.date(today.year - 1, 12, 31)
with col2:
    d = st.date_input(
        "Выберите период",
        (jan_1, datetime.date(today.year - 1, 12, 31)),
        jan_1,
        dec_31,
        format="DD.MM.YYYY",
    )

if d[0].strftime("%m") == '01':  # В идеале нужно объединить листы в один,
    month = 'Часы 01-23'         # но тогда я не уложился бы в срок :с
elif d[0].strftime("%m") == '02':
    month = 'Часы 02-23'
else:
    month = ''

if division == 'Подразделение1':
    df = pd.read_excel(df_all1, month)
elif division == 'Подразделение2':
    df = pd.read_excel(df_all2, month)
else:
    df = 0

start_date = int(d[0].strftime("%d")) + 2
try:
    if d[0].strftime("%m") == '01' and d[1].strftime("%m") == '02':
        st.text('К сожалению, такая выборка пока что не работает, поэтому выборка была ограничена последним днем первого месяца')
        end_date = 31 + 3
    else:
        end_date = int(d[1].strftime("%d")) + 3
except IndexError:
    end_date = start_date

df1 = df.iloc[:, list(range(start_date, end_date))]
indexes = list(range(start_date - 2, end_date - 2))
df1 = df1.count()
df1.index = indexes

st.bar_chart(df1)


# 2
col11, col12 = st.columns(2)
with col11:
    st.subheader('Соотношение долей работающих сотрудников (в процентах) разных подразделений за выбранный период:')

    df_pro1 = pd.read_excel(df_all1, month).iloc[:, list(range(start_date, end_date))]
    df_pro2 = pd.read_excel(df_all2, month).iloc[:, list(range(start_date, end_date))]
    df_pro = pd.DataFrame({
        'Подразделения': ['Подразделение1', 'Подразделение2'],
        'Сумма часов': [df_pro1.sum().sum(), df_pro2.sum().sum()],
    })
    pie = px.pie(df_pro, values='Сумма часов', names='Подразделения')
    st.plotly_chart(pie)



# 3
with col12:
    st.divider()
    d1 = st.date_input("Выберите день", datetime.date(2023, 1, 1))

    if d1.strftime("%m") == '01':
        month = 'Часы 01-23'
    elif d1.strftime("%m") == '02':
        month = 'Часы 02-23'
    else:
        month = ''

    df_count1 = pd.read_excel(df_all1, month).iloc[:, int(d1.strftime("%d")) + 2].count()
    df_count2 = pd.read_excel(df_all2, month).iloc[:, int(d1.strftime("%d")) + 2].count()

    st.subheader(f'Общее количество сотрудников в выбранный день: {df_count1 + df_count2}')
