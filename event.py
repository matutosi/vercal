import pandas as pd
from datetime import datetime, timedelta
import numpy as np

def generate_schedule(input_df):
    """
    入力データフレームから日程を生成し、除外日を除外する関数
    Args:
        input_df (pandas.DataFrame): 入力データフレーム
    Returns:
        pandas.DataFrame: 結果のデータフレーム
    """
    results = []
    for _, row in input_df.iterrows():
        dates_df = generate_dates(
            row["period_start"],
            row["period_end"],
            row["week_of_day"],
            row["event_start"],
            row["event_end"],
            row["event"],
        )
        except_list = row["except"].split(",")
        except_dates = [
            datetime.strptime(date, "%Y-%m-%d").date().strftime("%Y-%m-%d")
            for date in except_list
        ]
        dates_df = exclude_dates(dates_df, except_dates)
        results.append(dates_df)
    return pd.concat(results, ignore_index=True)

def generate_dates(period_start, period_end, week_of_day, event_start, event_end, event):
    """
    指定された期間と曜日から日程を抽出する関数
    """
    week_of_day_map = {"mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6}
    target_weekday = week_of_day_map[week_of_day.lower()]
    start_date = datetime.strptime(period_start, "%Y-%m-%d").date()
    end_date = datetime.strptime(period_end, "%Y-%m-%d").date()
    days_until_target_weekday = (target_weekday - start_date.weekday()) % 7
    current_date = start_date + timedelta(days=days_until_target_weekday)
    dates = []
    while current_date <= end_date:
        dates.append({
            "date": current_date.strftime("%Y-%m-%d"),
            "week_of_day": week_of_day,
            "event_start": event_start,
            "event_end": event_end,
            "event": event,
        })
        current_date += timedelta(days=7)
    return pd.DataFrame(dates)

def exclude_dates(df, except_dates):
    """
    指定された日付を除外する関数
    """
    return df[~df["date"].isin(except_dates)]

def format_events(df):
    """
    データフレームから日付とイベント情報を整形する関数
    Args:
        df (pandas.DataFrame): 入力データフレーム
    Returns:
        pandas.DataFrame: 整形後のデータフレーム
    """
    grouped = df.groupby("date").apply(
        lambda x: [create_event_dict(row) for _, row in x.iterrows()], include_groups=False).reset_index(name="event")
    return grouped

def create_event_dict(row):
    """
    行データからイベント辞書を作成する関数
    """
    event_dict = {"event_start": row["event_start"]}
    if not pd.isna(row["event_end"]):
        event_dict["event_end"] = row["event_end"]
    event_dict["event"] = row["event"]
    return event_dict

if __name__ == "__main__":
    # 使用例
    # input_data = {
    #     "period_start": ["2025-04-10", "2025-04-10"],
    #     "period_end": ["2025-07-10", "2025-07-10"],
    #     "week_of_day": ["wed", "wed"],
    #     "event_start": ["10:30", "12:30"],
    #     "event_end": ["12:00", np.nan],
    #     "event": ["math", "english"],
    #     "except": ["2025-05-05,2025-05-12", "2025-05-05,2025-05-# 12"],
    # }
    # input_df = pd.DataFrame(input_data)
    # input_df.to_excel(path)

    path = 'event.xlsx'
    input_df = pd.read_excel(path)
    date_df = generate_schedule(input_df)
    event_df = format_events(date_df)
