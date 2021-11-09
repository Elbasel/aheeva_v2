import pandas as pd
from datetime import datetime, timedelta
from xlrd import xldate_as_datetime
from matplotlib import cm
import os
import pickle


def floor_datetime(dt, delta):
    rounded = dt - (dt - datetime.min) % timedelta(minutes=delta)
    return rounded


def load_data(account):
    df = pd.read_csv(f'{account}_outbound.csv')
    df.rename(str.lower, axis='columns', inplace=True)
    df.rename(str.strip, axis='columns', inplace=True)
    df.rename(lambda x: x.replace(' ', '_'), axis='columns', inplace=True)
    df.start_time = df.start_time.apply(lambda x: xldate_as_datetime(float(x), 0) + timedelta(hours=2))
    df.end_time = df.end_time.apply(lambda x: xldate_as_datetime(float(x), 0) + timedelta(hours=2))
    df.duration = pd.to_timedelta(df.duration)

    return df


def get_outbound_status(df):
    if df['outbound_calls'] and df['status_shift'] == 'IN_CALL_OUTBOUND':
        return 1
    return 0


def get_total_talk_time(df):
    if df.outbound_status:
        return (df.duration + df.duration_shift).total_seconds()
    else:
        return 0


def get_outbound_calls(df):
     if df.status == 'DIALING' or df.status == 'ACCEPT_PREVIEW':
        return 1
     return 0



def center_text(s, props='text-align: center;'):
    return props


def strfdelta(tdelta, fmt):
    tdelta = timedelta(seconds=tdelta)
    d = {"days": str(tdelta.days)}
    d['hours'], rem = divmod(tdelta.seconds, 3600)
    d['hours'] = str(d['hours']).zfill(2)
    d['minutes'], d['seconds'] = divmod(rem, 60)
    d['minutes'] = str(d['minutes']).zfill(2)
    d['seconds'] = str(d['seconds']).zfill(2)
    return fmt.format(**d)


def analyze_outbound(account, start_date, end_date):
    
    print(f'loading {account} outbound data {start_date}_{end_date}...')

    df = load_data(account)
    if df.empty:
        return 0, 0, 0, 0, 0, 0, 0
    print('analyzing outbound data...')
    print('per interval...')
    lookup = pd.read_csv(f'{account}_lookup.csv')
    df['interval'] = df.start_time.apply(lambda x: pd.Timestamp.floor(x, f'30T').time())
    df['week_number'] = df.start_time.dt.strftime('%U').apply(lambda x: int(x) + 1)
    df['outbound_calls'] = df.apply(get_outbound_calls, axis=1)
    df = pd.merge(df, lookup, on='agent_id', how='left')
    df['date'] = df.start_time.apply(lambda x: x.date())
    df.to_csv('test.csv')


    df['status_shift'] = df.status.shift(-1)
    df['outbound_status'] = df.apply(get_outbound_status, axis=1)
    df['duration_shift'] = df.duration.shift(-1)
    df['total_talk_time'] = df.apply(get_total_talk_time, axis=1)
    df.dropna(axis=0, subset=['name'], inplace=True)
    print('outbound data loaded...')

    if int(df.outbound_calls.sum()) == 0:
        return 0, 0, 0, 0, 0
    interval_group = df.loc[:, ['interval', 'outbound_calls', 'outbound_status', 'total_talk_time']][df.outbound_calls==1].groupby('interval', as_index=False).sum()
    # interval_group.reset_index()
    # interval_group.drop(interval_group.index)
    interval_group['Average Handling Time'] = interval_group.total_talk_time / interval_group.outbound_status
    interval_group.fillna(0, inplace=True)
    interval_group['connection_rate'] = interval_group.outbound_status / interval_group.outbound_calls

    total_outbound_calls = int(interval_group.outbound_calls.sum())
    total_answered_calls = int(interval_group.outbound_status.sum())
    try:
        average_handling_time = str(timedelta(seconds=int(interval_group.total_talk_time.sum() / interval_group.outbound_status.sum())))
    except:
        average_handling_time = 'N/A'
    interval_group.drop('total_talk_time', axis='columns', inplace=True)
  
    
    interval_group.interval = interval_group.interval.apply(lambda x: str(x))
    bar_chart_df = interval_group.loc[:, ['interval', 'outbound_status', 'outbound_calls']].set_index('interval')
    bar_chart_df.rename(columns={'outbound_status': 'Answered', 'outbound_calls': ''}, inplace=True)
    interval_group.outbound_calls = interval_group.outbound_calls.apply(lambda x: str(x))
    interval_group.outbound_calls = interval_group.outbound_calls.apply(lambda x: str(x))
    interval_group.outbound_status = interval_group.outbound_status.apply(lambda x: str(x))
    interval_group.connection_rate = interval_group.connection_rate.apply(lambda x: f'{x:.1%}')

    interval_group.rename(columns={'interval': 'Interval', 'outbound_calls': 'Total Calls', 'outbound_status': 'Answered', 'connection_rate': 'Connection Rate', 'Average Handling Time': 'AHT'}, inplace=True)
    interval_group.set_index('Interval', inplace=True)
    interval_styler = interval_group.style.format(
        {   
            "AHT": lambda delta: strfdelta(delta, '{hours}:{minutes}:{seconds}'),
        }
    )
    interval_styler = interval_styler.applymap(center_text)

    cmap = cm.get_cmap('RdYlGn').reversed()
    interval_styler = interval_styler.background_gradient(subset='AHT', cmap=cmap)
    print('per agent...')

    agent_group = df.loc[:, ['name', 'outbound_calls', 'outbound_status', 'total_talk_time']].groupby('name', as_index=False).sum()




    agent_group['average_talk_time'] = agent_group.total_talk_time / agent_group.outbound_status
    agent_group.drop('total_talk_time', axis='columns', inplace=True)
    agent_group.fillna(0, inplace=True)

    agent_group['connection_rate'] = agent_group.outbound_status / agent_group.outbound_calls
    agent_group.connection_rate.fillna(0, inplace=True)

    agent_group.sort_values(by='average_talk_time', ascending=False, inplace=True)


    agent_group.outbound_calls = agent_group.outbound_calls.apply(lambda x: str(x))
    agent_group.outbound_calls = agent_group.outbound_calls.apply(lambda x: str(x))
    agent_group.outbound_status = agent_group.outbound_status.apply(lambda x: str(x))
    agent_group.connection_rate = agent_group.connection_rate.apply(lambda x: f'{x:.1%}')
    agent_group.rename(columns={'name': 'Name', 'outbound_calls': 'Total Calls', 'outbound_status': 'Answered', 'average_talk_time': 'AHT', 'connection_rate': 'Connection Rate'}, inplace=True)
    agent_group.set_index('Name', inplace=True)

    agent_styler = agent_group.style.format(
        {   
            "AHT": lambda delta: strfdelta(delta, '{hours}:{minutes}:{seconds}'),
            # 'connection_rate': lambda x: f'{x:.0%}',
        }
    )

    cmap = cm.get_cmap('RdYlGn').reversed()

    agent_styler.applymap(center_text)
    interval_styler.set_properties(**{'text-align': 'center'})
    agent_styler.set_properties(**{'text-align': 'center'})

    agent_styler = agent_styler.background_gradient(subset='AHT', cmap=cmap)
    print('per day...')

    date_group = df.loc[:, ['date', 'outbound_calls', 'outbound_status', 'total_talk_time']].groupby('date', as_index=False).sum()
    date_group['Average Handling Time'] = date_group.total_talk_time / date_group.outbound_status
    date_group['Average Handling Time'] = date_group['Average Handling Time'].apply(lambda x: strfdelta(x, '{hours}:{minutes}:{seconds}'))
    date_group.drop(['total_talk_time'], axis=1, inplace=True)
    date_group['connection_rate'] = date_group.outbound_status / date_group.outbound_calls
    
    date_group.connection_rate = date_group.connection_rate.apply(lambda x: f'{x:.1%}')
    date_group.outbound_calls = date_group.outbound_calls.apply(str)
    date_group.outbound_status = date_group.outbound_status.apply(str)
    date_group.date = date_group.date.apply(lambda x: x.strftime('%m/%d/%Y'))
    date_group.rename(columns={'outbound_calls': 'Total Calls', 'outbound_status': 'Answered', 'connection_rate': 'Connection Rate', 'date': 'Date', 'Average Handling Time': 'AHT'}, inplace=True)
    date_group.set_index('Date', inplace=True)
   
    date_styler = date_group.style
    date_styler.set_properties(**{'text-align': 'center'})
    
    print('outbound analysis over')
    return interval_styler, agent_styler, total_outbound_calls, total_answered_calls, average_handling_time, date_styler, bar_chart_df

