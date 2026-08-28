import datetime as dt
import os
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import event  # noqa: E402


def make_row(**kwargs):
    row = {
        'period_start': '2025-04-07',
        'period_end': '2025-04-28',
        'week_of_day': 'mon',
        'event_start': '12:30',
        'event_end': '13:00',
        'event': 'lunch',
        'except': np.nan,
    }
    row.update(kwargs)
    return pd.DataFrame([row])


class TestToDateStr:
    def test_str(self):
        assert event.to_date_str('2025-04-07') == '2025-04-07'

    def test_timestamp(self):
        # Excel の日付セルは Timestamp で読まれる
        assert event.to_date_str(pd.Timestamp('2025-04-07')) == '2025-04-07'

    def test_date(self):
        assert event.to_date_str(dt.date(2025, 4, 7)) == '2025-04-07'


class TestToTimeStr:
    def test_str(self):
        assert event.to_time_str('12:30') == '12:30'

    def test_time(self):
        # Excel の時刻セルは datetime.time で読まれる
        assert event.to_time_str(dt.time(12, 30)) == '12:30'

    def test_time_zero_minute(self):
        assert event.to_time_str(dt.time(9, 0)) == '9:00'


class TestGenerateDates:
    def test_weekly_dates(self):
        df = event.generate_dates('2025-04-07', '2025-04-28', 'mon', '12:30', '13:00', 'lunch')
        assert df['date'].tolist() == ['2025-04-07', '2025-04-14', '2025-04-21', '2025-04-28']

    def test_starts_from_first_target_weekday(self):
        # 期間の開始が水曜でも，最初の月曜から始まる
        df = event.generate_dates('2025-04-09', '2025-04-21', 'mon', '12:30', '13:00', 'lunch')
        assert df['date'].tolist() == ['2025-04-14', '2025-04-21']

    def test_week_of_day_is_case_insensitive(self):
        df = event.generate_dates('2025-04-07', '2025-04-07', 'MON', '12:30', '13:00', 'lunch')
        assert df['date'].tolist() == ['2025-04-07']

    def test_accepts_timestamp_period(self):
        df = event.generate_dates(pd.Timestamp('2025-04-07'), pd.Timestamp('2025-04-14'),
                                  'mon', '12:30', '13:00', 'lunch')
        assert df['date'].tolist() == ['2025-04-07', '2025-04-14']


class TestGenerateSchedule:
    def test_excludes_dates(self):
        df = event.generate_schedule(make_row(**{'except': '2025-04-14;2025-04-21'}))
        assert df['date'].tolist() == ['2025-04-07', '2025-04-28']

    def test_except_with_spaces(self):
        df = event.generate_schedule(make_row(**{'except': '2025-04-14; 2025-04-21'}))
        assert df['date'].tolist() == ['2025-04-07', '2025-04-28']

    def test_no_except(self):
        df = event.generate_schedule(make_row())
        assert len(df) == 4


class TestFormatEvents:
    def test_one_event_per_date(self):
        grouped = event.format_events(event.generate_schedule(make_row()))
        assert grouped['date'].tolist() == ['2025-04-07', '2025-04-14', '2025-04-21', '2025-04-28']
        assert grouped['event'].iloc[0] == [{'event_start': '12:30', 'event_end': '13:00', 'event': 'lunch'}]

    def test_missing_event_end_falls_back_to_start(self):
        # 終了時刻がないときは開始時刻と同じにする (描画側が event_end を必ず使うため)
        grouped = event.format_events(event.generate_schedule(make_row(event_end=np.nan)))
        assert grouped['event'].iloc[0] == [{'event_start': '12:30', 'event_end': '12:30', 'event': 'lunch'}]

    def test_time_cells_are_normalized(self):
        grouped = event.format_events(
            event.generate_schedule(make_row(event_start=dt.time(12, 30), event_end=dt.time(13, 0))))
        assert grouped['event'].iloc[0] == [{'event_start': '12:30', 'event_end': '13:00', 'event': 'lunch'}]

    def test_two_events_on_the_same_date(self):
        df = pd.concat([make_row(), make_row(event='meeting', event_start='15:00', event_end='16:00')],
                       ignore_index=True)
        grouped = event.format_events(event.generate_schedule(df))
        assert len(grouped['event'].iloc[0]) == 2


class TestSingleDayEvent:
    def test_empty_week_of_day_is_a_single_day(self):
        # 曜日が空の行は開始日の1日だけ (period_end は見ない)
        df = event.generate_schedule(make_row(week_of_day=np.nan, period_start='2025-04-10',
                                              period_end='2025-07-10'))
        assert df['date'].tolist() == ['2025-04-10']

    def test_empty_week_of_day_with_date_cell(self):
        df = event.generate_schedule(make_row(week_of_day=np.nan,
                                              period_start=pd.Timestamp('2025-04-10'),
                                              period_end=pd.Timestamp('2025-04-10')))
        assert df['date'].tolist() == ['2025-04-10']

    def test_no_event_start_is_an_all_day_event(self):
        # 時刻のない予定は終日の予定 (メモ欄に描かれる)
        grouped = event.format_events(
            event.generate_schedule(make_row(week_of_day=np.nan, event_start=np.nan,
                                             event_end=np.nan, event='入学式',
                                             period_start='2025-04-01', period_end='2025-04-01')))
        assert grouped['event'].iloc[0] == [{'event_start': None, 'event_end': None, 'event': '入学式'}]


class TestBundledTemplate:
    def test_shipped_schedule_xlsx_can_be_read(self):
        # 同梱の雛形をそのまま読めること (単発の予定の行を含む)
        here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        df_input = pd.read_excel(os.path.join(here, 'schedule.xlsx'))
        grouped = event.format_events(event.generate_schedule(df_input))
        assert len(grouped) > 0
        assert '2025-04-01' in grouped['date'].tolist()
