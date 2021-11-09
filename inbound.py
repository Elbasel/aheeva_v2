import os
import xlrd
import pandas as pd
import numpy as np
from matplotlib import cm
from datetime import datetime, timedelta
import pickle


def excel_datetime(date):
    return xlrd.xldate_as_datetime(date, 0)


def strfdelta(tdelta, fmt):
    d = {"days": str(tdelta.days)}
    d['hours'], rem = divmod(tdelta.seconds, 3600)
    d['hours'] = str(d['hours']).zfill(2)
    d['minutes'], d['seconds'] = divmod(rem, 60)
    d['minutes'] = str(d['minutes']).zfill(2)
    d['seconds'] = str(d['seconds']).zfill(2)
    return fmt.format(**d)


def dt_to_td(dt):
    dt = dt.time()
    hours = dt.hour
    minutes = dt.minute
    seconds = dt.second
    
    total_seconds = hours * 60 * 60 + minutes * 60 + seconds
    return timedelta(seconds=total_seconds)


def center_text(s, props='text-align: center;'):
    return props


def analyze_inbound(account, start_date, end_date):

    print(f'loading {account} inbound data {start_date}_{end_date}')
    df = pd.read_csv(f'{account}_inbound.csv')
    queues = pd.read_csv(f'{account}_queues.csv')
    lookup = pd.read_csv(f'{account}_lookup.csv')
    if account == 'swvl':
        lookup = pd.read_csv('swvl_inbound_lookup.csv')
    df['Event Time'] = pd.to_datetime(df['Event Time']) + timedelta(hours=2)
    df['interval'] = df['Event Time'].apply(lambda x: pd.Timestamp.floor(x, f'30T').time())
    df.Duration = df.Duration.apply(excel_datetime)
    df['Talk time'] = df['Talk time'].apply(excel_datetime)
    df.Duration = df.Duration.apply(dt_to_td)
    df['Talk time'] = df['Talk time'].apply(dt_to_td)
    df['total_talk_time'] = df['Duration'] + df['Talk time']
    df = pd.merge(df, queues, on='Queue', how='left')
    df.rename(columns={'Agent': 'agent_id'}, inplace=True)
    df.dropna(subset=['offered'], inplace=True)
    df = pd.merge(df, lookup, on='agent_id', how='left')
    df['Abandoned'] = df['Talk time'].apply(lambda x: 1 if x == timedelta(0, 0, 0) else 0)
    df['answered_within'] = df[df.Abandoned == 0].Duration.apply(lambda x: 1 if x <= timedelta(0, 20, 0) else 0)
    df['date'] = df['Event Time'].apply(lambda x: x.date())
    print('inbound data loaded')
    print('analyzing inbound data')
    print('per interval')
    interval_group = df.loc[:, ['interval','offered', 'total_talk_time', 'Abandoned', 'answered_within']].groupby('interval', as_index=False).agg({'offered': np.sum, 'Abandoned': np.sum, 'total_talk_time': np.sum, 'answered_within': np.sum})
    
    
    interval_group['Answered'] = interval_group.offered - interval_group.Abandoned
    interval_group['average_handling_time'] = interval_group.total_talk_time / interval_group.offered
    interval_group = interval_group[['interval', 'offered', 'Answered', 'answered_within', 'Abandoned', 'average_handling_time', 'total_talk_time']]
    
    
    total_inbound_calls = int(interval_group.offered.sum())
    total_answered_calls = int(interval_group.Answered.sum())
    total_abandoned_calls = total_inbound_calls - total_answered_calls
    # average_handling_time = str(interval_group.total_talk_time.sum() / interval_group.Answered.sum())

    try:
        average_handling_time = strfdelta(interval_group.total_talk_time.sum() / total_answered_calls, '{minutes}:{seconds}')
    except ZeroDivisionError:
        average_handling_time = 'N/A'
    interval_group.drop(['total_talk_time'], axis=1, inplace=True)
    interval_group.average_handling_time = interval_group.average_handling_time.apply(lambda delta: strfdelta(delta, '{hours}:{minutes}:{seconds}'))
    interval_group.interval = interval_group.interval.apply(lambda x: str(x))
    
    
    
    interval_group['SLA'] = interval_group.answered_within / interval_group.offered
    interval_group.SLA = interval_group.SLA.apply(lambda x: f'{x:.0%}')
    interval_group.offered = interval_group.offered.apply(lambda x: str(int(x)))
    interval_group.Answered = interval_group.Answered.apply(lambda x: str(int(x)))
    try:
        interval_group.answered_within = interval_group.answered_within.apply(lambda x: str(int(x)))
    except TypeError:
        interval_group.answered_within = '0'
        average_handling_time = 'N/A'


    interval_group.Abandoned = interval_group.Abandoned.apply(lambda x: str(int(x)))

    interval_group.rename(columns={'interval': 'Interval', 'offered': 'Offered', 'answered_within': 'Answred Within', 'average_handling_time': 'AHT'}, inplace=True)
    interval_group = interval_group.loc[:, ['Interval', 'Offered', 'Answered', 'Answred Within', 'SLA', 'AHT']]

    # interval_group.reset_index()
    interval_group.set_index('Interval', inplace=True)
    # interval_group.index = interval_group.Interval
    
    interval_styler = interval_group.style #.format(
    #     {   
    #         "Average Handling Time": lambda delta: strfdelta(delta, '{hours}:{minutes}:{seconds}'),
    #     }
    # )
    interval_styler.applymap(center_text)
    interval_styler.set_properties(**{'text-align': 'center'})
    print('per agent')

    agent_group = df.loc[:, ['name','offered', 'total_talk_time', 'Abandoned', 'answered_within']].groupby('name', as_index=False).agg({'offered': np.sum, 'Abandoned': np.sum, 'total_talk_time': np.sum, 'answered_within': np.sum})
    agent_group['Answered'] = agent_group['offered'] - agent_group.Abandoned
    agent_group['average_handling_time'] = agent_group.total_talk_time / agent_group.offered
    agent_group.drop(['total_talk_time'], axis=1, inplace=True)
    agent_group = agent_group[['name', 'offered', 'Answered', 'answered_within', 'Abandoned', 'average_handling_time']]
    agent_group['SLA'] = (agent_group.answered_within / agent_group.offered).apply(lambda x: f'{x:.0%}')

    agent_group.average_handling_time = agent_group.average_handling_time.apply(lambda delta: strfdelta(delta, '{hours}:{minutes}:{seconds}'))
    agent_group.offered = agent_group.offered.apply(lambda x: str(int(x)))
    agent_group.Answered = agent_group.Answered.apply(lambda x: str(int(x)))
    agent_group.answered_within = agent_group.answered_within.apply(lambda x: str(int(x)))
    agent_group.Abandoned = agent_group.Abandoned.apply(lambda x: str(int(x)))
    agent_group.rename(columns={'name': 'Name', 'offered': 'Offered', 'answered_within': 'Answred Within', 'average_handling_time': 'AHT'}, inplace=True)
    agent_group.dropna(subset=['Name'])
    agent_group.set_index('Name', inplace=True)
    agent_group.drop(['Offered', 'Abandoned'], axis=1, inplace=True)
    agent_styler = agent_group.style 
    agent_styler.applymap(center_text)
    agent_styler.set_properties(**{'text-align': 'center'})
    print('per day')

    date_group = df.loc[:, ['date', 'offered', 'Abandoned', 'answered_within', 'total_talk_time']].groupby('date', as_index=False).agg({'total_talk_time': np.sum, 'offered': np.sum, 'Abandoned': np.sum, 'answered_within': np.sum}) 
    date_group.total_talk_time = date_group.total_talk_time.apply(lambda x: x.total_seconds())
    date_group['Answered'] = date_group.offered - date_group.Abandoned
    date_group['SLA'] = date_group.answered_within / date_group.offered
    date_group = date_group.loc[:, ['date', 'offered', 'Answered', 'answered_within', 'Abandoned', 'SLA', 'total_talk_time']]
    date_group['AHT'] = date_group.total_talk_time / date_group.Answered
    
    total_sla = date_group.answered_within.sum() / date_group.offered.sum()
    total_sla = f'{total_sla:.0%}'
    total_anwered_within = int(date_group.answered_within.sum())
    
    date_group.drop(['total_talk_time'], axis=1, inplace=True)
    date_group.offered = date_group.offered.apply(lambda x: str(int(x)))
    date_group.Answered = date_group.Answered.apply(lambda x: str(int(x)))
    date_group.answered_within = date_group.answered_within.apply(lambda x: str(int(x)))
    date_group.Abandoned = date_group.Abandoned.apply(lambda x: str(int(x)))
    date_group.SLA = date_group.SLA.apply(lambda x: f'{x:.0%}')
    date_group.replace([np.inf, -np.inf], 0, inplace=True)
    date_group.AHT = date_group.AHT.apply(lambda x: strfdelta(timedelta(seconds=int(x)), '{hours}:{minutes}:{seconds}'))
    date_group.set_index('date', inplace=True)
    date_group.rename(columns={'offered': 'Total', 'answered_within': 'Within SLA', }, inplace=True)
    date_styler = date_group.style

    date_styler.applymap(center_text)
    print('inbound analysis over')

    return interval_styler, agent_styler, total_inbound_calls, total_answered_calls, total_abandoned_calls, average_handling_time, date_styler, total_sla, total_anwered_within


