import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title='Data', layout='wide')

st.title('Data Analysis')

excel_file='data.xlsx'
sheet_name1='Filter'
sheet_name2='Service'

df0=pd.read_excel(excel_file, sheet_name=sheet_name2, usecols='C:D', header=1, nrows=12)
df1=pd.read_excel(excel_file, sheet_name=sheet_name2, usecols='G:H', header=1, nrows=12)
df2=pd.read_excel(excel_file, sheet_name=sheet_name2, usecols='K:L', header=1, nrows=12)
df3=pd.read_excel(excel_file, sheet_name=sheet_name1, usecols='E:G', header=1, nrows=12)
df4=pd.read_excel(excel_file, sheet_name=sheet_name1, usecols='J:L', header=1, nrows=5)

st.markdown('##')

### --- KPI
quantity=int(df0["Quantity"].sum())
net=int(df1["Net"].sum())
workh=int(df2["Work"].sum())

col1, col2, col3, col4, col5 =st.columns(5)

col2.metric(label="Quantity", value=(f'{quantity:,}'), delta=1000)
col3.metric(label="Net", value=(f'€{net:,}'), delta=1000)
col4.metric(label="Work Hour", value=(f'{workh:,}'), delta=1000)

bar_chart = px.bar(df0, x='Quantity', y='Service', title='Quantity', orientation="h", color_discrete_sequence=["#ff053a"]*len(df0), template='plotly_white')

bar_chart1 = px.bar(df1, x='Net', y='Service1', title='Net', orientation="h", color_discrete_sequence=["#ff053a"]*len(df1), template='plotly_white')

pie_chart = px.pie(df4, names='Service3', values='Net.1', title='Best Net Value', template='plotly_white')

col6, col7 =st.columns(2)

st.markdown('##')
st.markdown('##')

col6.plotly_chart(bar_chart, use_container_width=True)
col7.plotly_chart(bar_chart1, use_container_width=True)

st.markdown('##')

col8, col9, col10 = st.columns(3)

with col8:
    st.dataframe(df1)

col10.plotly_chart(pie_chart, use_container_width=True)

with col9:
    st.dataframe(df2)


## END