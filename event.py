from datetime import datetime, timedelta
import numpy as np
import pandas as pd

def to_date_str(value):
    """
    normalize a date cell into 'YYYY-MM-DD'
    Excel の日付セルは Timestamp で読まれるので，文字列と両方を受ける
    """
    if isinstance(value, str):
        return datetime.strptime(value, '%Y-%m-%d').date().strftime('%Y-%m-%d')
    return pd.Timestamp(value).date().strftime('%Y-%m-%d')

def to_time_str(value):
    """
    normalize a time cell into 'H:MM'
    Excel の時刻セルは datetime.time で読まれるので，文字列と両方を受ける
    """
    if isinstance(value, str):
        return value
    return f'{value.hour}:{value.minute:02d}'

def generate_schedule(input_df):
    """
    generate schedule from dataframe
    Args:
        input_df (pandas.DataFrame)
    Returns:
        pandas.DataFrame
    """
    results = []
    for _, row in input_df.iterrows():
        dates_df = generate_dates(
            row['period_start'],
            row['period_end'],
            row['week_of_day'],
            row['event_start'],
            row['event_end'],
            row['event'],
        )
        if isinstance(row['except'], str):
            except_list = row['except'].split(';')
        else:
            except_list = []
        except_dates = [to_date_str(date.strip()) for date in except_list]
        dates_df = exclude_dates(dates_df, except_dates)
        results.append(dates_df)
    return pd.concat(results, ignore_index=True)

def generate_dates(period_start, period_end, week_of_day, event_start, event_end, event):
    """
    generate dates within input period
    """
    week_of_day_map = {'mon': 0, 'tue': 1, 'wed': 2, 'thu': 3, 'fri': 4, 'sat': 5, 'sun': 6}
    start_date = datetime.strptime(to_date_str(period_start), '%Y-%m-%d').date()
    end_date = datetime.strptime(to_date_str(period_end), '%Y-%m-%d').date()
    if pd.isna(week_of_day):
        # 曜日が空の行は単発の予定として，開始日の1日だけに置く
        end_date = start_date
        current_date = start_date
        step = timedelta(days=1)
    else:
        target_weekday = week_of_day_map[week_of_day.lower()]
        days_until_target_weekday = (target_weekday - start_date.weekday()) % 7
        current_date = start_date + timedelta(days=days_until_target_weekday)
        step = timedelta(days=7)
    dates = []
    while current_date <= end_date:
        dates.append({
            'date': current_date.strftime('%Y-%m-%d'),
            'week_of_day': week_of_day,
            'event_start': event_start,
            'event_end': event_end,
            'event': event,
        })
        current_date += step
    return pd.DataFrame(dates)

def exclude_dates(df, except_dates):
    """
    exclude dates
    """
    return df[~df['date'].isin(except_dates)]

def format_events(df):
    """
    from dataframe into event format 
    Args:
        df (pandas.DataFrame)
    Returns:
        pandas.DataFrame
    """
    grouped = df.groupby('date').apply(
        lambda x: [create_event_dict(row) for _, row in x.iterrows()], include_groups=False).reset_index(name='event')
    return grouped

def create_event_dict(row):
    """
    create event dict from dataframe row
    """
    if pd.isna(row['event_start']):
        # 開始時刻がないものは終日の予定 (時間軸に置けないのでメモ欄へ描く)
        return {'event_start': None, 'event_end': None, 'event': row['event']}
    event_start = to_time_str(row['event_start'])
    # 終了時刻がないときは開始時刻と同じにする (線1本で描かれる)
    event_end = event_start if pd.isna(row['event_end']) else to_time_str(row['event_end'])
    return {'event_start': event_start, 'event_end': event_end, 'event': row['event']}
