import os
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import vercal  # noqa: E402


class TestStringToFloat:
    def test_half_hour(self):
        assert vercal.string2float('9:30') == 9.5

    def test_zero_padded(self):
        assert vercal.string2float('09:00') == 9.0


class TestUseFont:
    def test_basename_without_extension(self):
        assert vercal.use_font('./fonts/HackGen35Console-Regular.ttf') == 'HackGen35Console-Regular'


class TestCreateYear:
    def test_april_start_begins_with_april(self):
        df = vercal.create_year(2025)
        first, last = df.iloc[0], df.iloc[-1]
        assert (first['year'], first['month'], first['day']) == (2025, 4, 1)
        assert (last['year'], last['month'], last['day']) == (2026, 3, 31)

    def test_january_start_begins_with_january(self):
        df = vercal.create_year(2025, start_april=False)
        first, last = df.iloc[0], df.iloc[-1]
        assert (first['year'], first['month'], first['day']) == (2025, 1, 1)
        assert (last['year'], last['month'], last['day']) == (2025, 12, 31)

    def test_leap_year(self):
        # 2024-04 始まりは 2025-02-28 までなので 365 日，2023-04 始まりは閏日を含み 366 日
        assert len(vercal.create_year(2023)) == 366
        assert len(vercal.create_year(2024)) == 365


class TestAddPosition:
    def test_monday_start_adjust_left(self):
        df = vercal.add_position(pd.DataFrame({'weekday': ['mon', 'thu', 'fri', 'sun']}))
        assert df['position'].tolist() == [0, 3, 0, 2]

    def test_sunday_start(self):
        df = vercal.add_position(pd.DataFrame({'weekday': ['sun', 'mon']}), starts_with_mon=False)
        assert df['position'].tolist() == [0, 1]

    def test_adjust_right(self):
        df = vercal.add_position(pd.DataFrame({'weekday': ['mon', 'thu']}), adjust_left=False)
        assert df['position'].tolist() == [1, 0]


class TestAddPage:
    def test_page_increments_when_position_goes_back(self):
        df = vercal.add_page(pd.DataFrame({'position': [1, 2, 3, 0, 1, 2, 0]}))
        assert df['page'].tolist() == [1, 1, 1, 2, 2, 2, 3]


    def test_works_with_a_non_range_index(self):
        # concat の後など，添字が 0 から並んでいなくても同じ結果になる
        df = pd.DataFrame({'position': [1, 2, 3, 0, 1]}, index=[10, 11, 12, 13, 14])
        assert vercal.add_page(df)['page'].tolist() == [1, 1, 1, 2, 2]


class TestAddDrawYearMonth:
    def test_first_day_of_month_and_first_of_page(self):
        df = pd.DataFrame({'day': [30, 1, 2, 3], 'position': [0, 1, 2, 0]})
        df = vercal.add_draw_year_month(df)
        assert df['draw_year_month'].tolist() == [False, True, False, True]


class TestCreateYearDf:
    def test_columns_and_length(self):
        df = vercal.create_year_df(2025)
        assert list(df.columns) == ['year', 'month', 'day', 'weekday', 'position', 'page', 'draw_year_month']
        assert len(df) == 365

    def test_four_days_per_page_at_most(self):
        df = vercal.create_year_df(2025)
        assert df.groupby('page').size().max() <= 4


class TestCalendarWeeklyVertical:
    def test_rejects_empty_day_range(self):
        with pytest.raises(ValueError):
            vercal.calendar_weekly_vertical(2025, hour_start=6, hour_end=6)

    def test_creates_pdf(self, tmp_path):
        path = str(tmp_path / 'calendar.pdf')
        out = vercal.calendar_weekly_vertical(
            2025, calendar_path=path, font_path='./HackGen35Console-Regular.ttf',
            hour_start=6, hour_end=22)
        assert out == path
        assert os.path.getsize(path) > 0
        with open(path, 'rb') as f:
            assert f.read(5) == b'%PDF-'

    def test_creates_pdf_with_events(self, tmp_path):
        import event
        df_input = pd.DataFrame([{
            'period_start': '2025-04-07', 'period_end': '2025-04-28', 'week_of_day': 'mon',
            'event_start': '12:30', 'event_end': '13:00', 'event': 'lunch', 'except': None}])
        df_event = event.format_events(event.generate_schedule(df_input))
        path = str(tmp_path / 'calendar_event.pdf')
        vercal.calendar_weekly_vertical(
            2025, calendar_path=path, font_path='./HackGen35Console-Regular.ttf',
            hour_start=6, hour_end=22, df_event=df_event)
        assert os.path.getsize(path) > 0


class TestWrapText:
    FONT = 'HackGen35Console-Regular'

    @classmethod
    def setup_class(cls):
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        pdfmetrics.registerFont(TTFont(cls.FONT, os.path.join(here, 'HackGen35Console-Regular.ttf')))

    def width(self, text, size=8):
        from reportlab.pdfbase import pdfmetrics
        return pdfmetrics.stringWidth(text, self.FONT, size)

    def test_short_text_is_one_line(self):
        assert vercal.wrap_text('math', self.FONT, 8, self.width('mathmath')) == ['math']

    def test_long_text_is_wrapped(self):
        max_width = self.width('12345')
        lines = vercal.wrap_text('from 15:00 to 17:00', self.FONT, 8, max_width)
        assert len(lines) > 1
        assert ''.join(lines) == 'from 15:00 to 17:00'
        assert all(self.width(line) <= max_width for line in lines)

    def test_japanese_is_wrapped_by_character(self):
        max_width = self.width('あいう')
        lines = vercal.wrap_text('あいうえおかきくけこ', self.FONT, 8, max_width)
        assert ''.join(lines) == 'あいうえおかきくけこ'
        assert all(self.width(line) <= max_width for line in lines)

    def test_max_lines_truncates_with_ellipsis(self):
        lines = vercal.wrap_text('from 15:00 to 17:00', self.FONT, 8, self.width('12345'), max_lines=1)
        assert len(lines) == 1
        assert lines[0].endswith('...')
        assert self.width(lines[0]) <= self.width('12345')

    def test_empty_text(self):
        assert vercal.wrap_text('', self.FONT, 8, 100) == ['']


class TestAllDayEvent:
    def test_creates_pdf_from_the_bundled_template(self, tmp_path):
        import event
        here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        df_input = pd.read_excel(os.path.join(here, 'schedule.xlsx'))
        df_event = event.format_events(event.generate_schedule(df_input))
        path = str(tmp_path / 'calendar_template.pdf')
        vercal.calendar_weekly_vertical(
            2025, calendar_path=path, font_path=os.path.join(here, 'HackGen35Console-Regular.ttf'),
            hour_start=6, hour_end=22, df_event=df_event)
        assert os.path.getsize(path) > 0


class TestCli:
    def test_creates_pdf_without_schedule(self, tmp_path):
        path = str(tmp_path / 'cli.pdf')
        out = vercal.main(['--year', '2025', '--out', path, '--hour-end', '22'])
        assert out == path
        assert os.path.getsize(path) > 0

    def test_creates_pdf_with_schedule(self, tmp_path):
        here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = str(tmp_path / 'cli_schedule.pdf')
        vercal.main(['--year', '2025', '--out', path,
                     '--schedule', os.path.join(here, 'schedule.xlsx')])
        assert os.path.getsize(path) > 0

    def test_default_font_is_bundled(self):
        assert os.path.exists(vercal.DEFAULT_FONT_PATH)
