import streamlit as st
import pandas as pd
import numpy as np
from xlrd import xldate_as_datetime
from time import sleep
from datetime import datetime, timedelta
from matplotlib import cm
import os
import time
import pickle

from get_data import update_data
from outbound import analyze_outbound
from inbound import analyze_inbound

from outbound import load_data



def current_interval():
    now = datetime.now()
    rounded = now - (now - datetime.min) % timedelta(minutes=30)
    return rounded

def refresh_data(account, start_date, end_date, interval=current_interval()):
    global last_account
            
    with open('last_refresh.pickle', 'rb') as last_refersh:
        last_account, last_start_date, last_end_date, last_interval = pickle.load(last_refersh)
    if any([last_account != account, last_start_date != start_date, last_end_date != end_date, last_interval != interval]):
        update_data(account, start_date, end_date)
        with open('last_refresh.pickle', 'wb') as last_referesh:
            pickle.dump([account, start_date, end_date, interval], last_referesh)

st.set_page_config(page_title='Real-Time')
sleep(3)
os.system('clear')

TODAY = datetime.today()
with open('last_refresh.pickle', 'rb') as last_refersh:
        last_account, last_start_date, last_end_date, last_interval = pickle.load(last_refersh)

# start_date = end_date = TODAY.date()
account_list = ['Furless', 'SWVL', 'BAT', 'Rizkalla', 'Udacity', 'ZoomCar']
account_index =  list(map(str.lower, account_list)).index(last_account) if last_account else 1
# account = st.sidebar.selectbox('Account', account_list, index=account_index).lower()
account = 'swvl'
start_date = st.sidebar.date_input('From')
end_date = st.sidebar.date_input('To')



if end_date - start_date < timedelta(0, 0, 0):
    st.header('Please choose a correct date interval')
    st.stop()

refresh_data(account, start_date, end_date)

refresh_button = st.sidebar.button('Refresh', on_click=update_data, args=[account.lower(), start_date, end_date])
try:
    with open(f'{account}_outbound.csv') as outbound_file:
        st.sidebar.download_button('Download Outbound Raw Data', outbound_file, f'{account} Outbound {start_date} to {end_date}.csv')

    with open(f'{account}_inbound.csv') as inbound_file:
        st.sidebar.download_button('Download Inbound Raw Data', inbound_file, f'{account} Inbound {start_date} to {end_date}.csv')
except FileNotFoundError:
    if datetime.now().hour < 8:

        st.header('Everyone is sleeping right now!')
        st.header('There\'s no data to display for today')
        st.header('Try changing the date on the right to the previous day to see results')
        st.stop(_)
    else:
        st.header('Bad Credentials ERROR 402')
        st.stop()

# st.title('Real Time Performance')

st.title(f'{account.capitalize()} Egypt Performance')
if start_date == end_date: 
    st.header(f'{start_date.strftime("%B %-d")}')
else:
    st.header(f'{start_date.strftime("%B")} {start_date.day} To {end_date.strftime("%B")} {end_date.day}')
st.caption('Made with ❤️ by Abdulrahman Elbasel For Centro CDX')


outbound_header = 'Outbound' if account != 'swvl' else 'Customer Experience'
inbound_header = 'Inbound' if account != 'swvl' else 'Captain Helpline'
try:
    outbound_interval_styler, outbound_agent_styler, total_outbound_calls, answered_calls, total_aht, outbound_date_group, outbound_bar_chart = analyze_outbound(account, start_date, end_date)
except ValueError:
    st.header('Shift Hasn\'t started yet')
    st.write('Please select a different date interval from the left sidebar')
    st.stop()

if account != 'udacity':
    inbound_interval_styler, inbound_agent_styler, total_inbound_calls, total_answered_calls, total_abandoned_calls, inbound_aht, inbound_date_styler, total_sla, total_answered_within  = analyze_inbound(account, start_date, end_date)

st.header(outbound_header)

if outbound_interval_styler != 0:
   
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric('Total Calls', total_outbound_calls)

    with col2:
        st.metric('Answered', answered_calls)

    with col3:
        st.metric('Connection Rate', f'{answered_calls/total_outbound_calls:.1%}')

    with col4:
        st.metric('AHT', total_aht)
    st.write(outbound_date_group)
    
    st.write(outbound_interval_styler)
    st.write(outbound_agent_styler)
    st.bar_chart(outbound_bar_chart)
else:
    st.write('No Outbound Calls')
if account != 'udacity':
    st.header(inbound_header)

    col1_in, col2_in, col3_in, col4_in, col5_in, col6_in = st.columns(6)
    with col1_in:
        st.metric('Total Calls', total_inbound_calls)

    with col3_in:
        st.metric('Ans within', total_answered_within)

    with col2_in:
        st.metric('Ans', total_answered_calls)

    with col4_in:
        st.metric('Abandoned', total_abandoned_calls)
    
    with col5_in:
        st.metric('SLA', total_sla)

    with col6_in:
        st.metric('AHT', inbound_aht)
    

    st.write(inbound_date_styler)
    st.write(inbound_interval_styler)
    # if account != 'swvl':
    st.write(inbound_agent_styler)


st.caption('')
st.caption('')
st.caption('')
st.caption('')





print('Dashboard is running')
print('Last Updated:', datetime.today())
# print('Elbasel©')

