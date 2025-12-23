import os
import calendar
from datetime import datetime
import pandas as pd

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5
from reportlab.lib.units import mm

# --- データ作成系関数 ---
def create_year_df(year, 
                   start_april=True, starts_with_mon=True, adjust_left=True):
    df_year = create_year(year, start_april=start_april)
    df_year = add_position(df_year, starts_with_mon=starts_with_mon, adjust_left=adjust_left)
    df_year = add_page(df_year)
    df_year = add_draw_year_month(df_year)
    return df_year

def create_year(year, start_april=True):
    if not start_april:
        return generate_dates(year)
    cal_this_yr = generate_dates(year).query('month > 3')
    cal_next_yr = generate_dates(year + 1).query('month < 4')
    return pd.concat([cal_this_yr, cal_next_yr])

def generate_dates(year):
    data = []
    for month in range(1, 13):
        for day in range(1, calendar.monthrange(year, month)[1] + 1):
            weekday_num = calendar.weekday(year, month, day)
            weekday_abbr = calendar.day_abbr[weekday_num].lower()
            data.append([year, month, day, weekday_abbr])
    return pd.DataFrame(data, columns=['year', 'month', 'day', 'weekday'])

def add_page(df, base_col='position', page_col='page'):
    df[page_col] = 1
    for i in range(1, len(df)):
        if df[base_col].iloc[i] < df[base_col].iloc[i - 1]:
            df.loc[i:, page_col] += 1
    return df

def add_position(df, starts_with_mon=True, adjust_left=True):
    wday_order = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'] if starts_with_mon else ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat']
    pos_map = [0, 1, 2, 3, 0, 1, 2] if adjust_left else [1, 2, 3, 0, 1, 2, 3]
    wday = pd.DataFrame({'weekday': wday_order, 'position': pos_map})
    return df.merge(wday, how='left')

def add_draw_year_month(df):
    df['draw_year_month'] = False
    df.loc[df['day'] == 1, 'draw_year_month'] = True
    for i in range(1, len(df)):
        if df['position'].iloc[i] < df['position'].iloc[i - 1]:
            df.loc[i, 'draw_year_month'] = True
    return df

# --- 描画系補助関数 ---
def string2float(time_str):
    hours, minutes = map(int, time_str.split(':'))
    return hours + minutes / 60.0

def use_font(font_path):
    return os.path.splitext(os.path.basename(font_path))[0]

# --- セクション描画関数 ---
def date_section(c, left, top, year, month, day, wday, 
                 h_year_month, h_wday_day, draw_year_month=True):
    if draw_year_month:
        c.drawString(left, top - c._fontsize, f'{year}-{month.zfill(2)}')
    c.drawString(left, top - h_wday_day - c._fontsize, f'{day.zfill(2)} {wday}')

def memo_section(c, left, right, top, h_year_month, h_wday_day, h_memo, circle_no=3):
    circle_distance = h_memo / circle_no
    top_memo = top - (h_year_month + h_wday_day) - circle_distance * 0.5 
    c.setDash([1, 0])
    c.setLineWidth(0.8)
    top_line = top - (h_year_month + h_wday_day)
    bottom_line = top - (h_year_month + h_wday_day + h_memo)
    c.line(left, top_line, right, top_line)
    c.line(left, bottom_line, right, bottom_line)
    for i in range(circle_no):
        c.circle(left + circle_distance * 0.5, top_memo - circle_distance * i, circle_distance / 3)

def ten_minute(c, left, right, top, h_hour, hour_start, hour_end):
    hours = hour_end - hour_start
    one_hour = h_hour / hours
    ten_minutes = one_hour / 6
    x = left - (left - right) * 0.12
    c.setDash([1, 0])
    c.setLineWidth(1)
    c.setStrokeColorRGB(0, 0.7, 0.3)
    for i in range((hours) * 6):
        y = top - ten_minutes * i
        circle_size = ((i % 2) + 1) * 0.5
        c.circle(x, y, circle_size)
    c.setStrokeColorRGB(0, 0, 0)

def hour_section(c, left, right, top_hour, h_hour, 
                 hour_start, hour_end, font_size_hour, draw_numbers=True):
    c.setDash([3, 2])
    c.setLineWidth(0.5)
    hours = hour_end - hour_start
    one_hour = h_hour / hours
    if draw_numbers:
        c.setFont(c._fontname, font_size_hour)
        for i in range(hours):
            c.drawString(left, top_hour - one_hour * i - font_size_hour, str(hour_start + i).zfill(2))
    for i in range(1, hours + 1):
        c.line(left, top_hour - one_hour * i, right, top_hour - one_hour * i)

def draw_schedule(c, schedule, event_start, event_end, 
                  hour_start, hour_end, top_hour, 
                  left, width, h_hour, font_size_hour):
    hours = hour_end - hour_start
    one_hour = h_hour / hours
    x = left
    y = top_hour - (string2float(event_start) - hour_start) * one_hour
    duration = - (string2float(event_end) - string2float(event_start)) * one_hour
    c.setDash([1, 0])
    c.setLineWidth(0.5)
    c.rect(x + width * 0.12, y, width * 0.83, duration)
    c.setFont(c._fontname, font_size_hour)
    c.drawString(x + width * 0.13, y - font_size_hour, schedule)

# --- ブロック描画の統合 ---
def draw_common_skeleton(c, left, right, top, height, 
                         hour_start, hour_end, h_year_month, h_wday_day, h_memo):
    """日付あり・なし共通の骨組み（メモ欄の線と時間線）を描画"""
    top_hour = top - (h_year_month + h_wday_day + h_memo)
    h_hour = top_hour - (top - height)
    # メモセクションの描画（共通）
    memo_section(c, left, right, top, h_year_month, h_wday_day, h_memo)
    return top_hour, h_hour

def create_day(c, left, top, width, height, 
               font_name, font_size, year, month, day, wday, 
               hour_start=6, hour_end=24, h_year_month=5*mm, h_wday_day=5*mm, h_memo=15*mm,
               draw_year_month=True, df_event=pd.DataFrame()):
    """通常の日付ブロック（日付・10分丸・時間数字あり）"""
    offset = 1 * mm
    right = left + width - offset
    c.setFont(font_name, font_size)
    font_size_hour = font_size * 0.7
    # 共通の骨組み
    top_hour, h_hour = draw_common_skeleton(c, left, right, top, height, hour_start, hour_end, h_year_month, h_wday_day, h_memo)
    # 日付あり特有：10分間隔の丸を描画
    ten_minute(c, left, right, top_hour, h_hour, hour_start, hour_end)
    # 日付セクションと時間数字の描画
    date_section(c, left, top, year, month, day, wday, h_year_month, h_wday_day, draw_year_month)
    hour_section(c, left, right, top_hour, h_hour, hour_start, hour_end, font_size_hour, draw_numbers=True)
    # イベントの描画
    date_today = f"{int(year)}-{int(month):02d}-{int(day):02d}"
    events = df_event[df_event['date'] == date_today]
    if not events.empty:
        for _, event_row in events.iterrows():
            if isinstance(event_row['event'], list):
                for event in event_row['event']:
                    draw_schedule(c, event['event'], event['event_start'], event['event_end'], 
                                  hour_start, hour_end, top_hour, left, width, h_hour, font_size_hour)

def draw_empty_block(c, left, top, width, height, 
                     hour_start, hour_end, h_year_month=5*mm, h_wday_day=5*mm, h_memo=15*mm):
    """日付のない空ブロック（メモ欄の線と時間線のみ）"""
    offset = 1 * mm
    right = left + width - offset
    # 共通の骨組み
    top_hour, h_hour = draw_common_skeleton(c, left, right, top, height, hour_start, hour_end, h_year_month, h_wday_day, h_memo)
    # 10分丸は描かず、時間線のみ描画（数字なし）
    hour_section(c, left, right, top_hour, h_hour, hour_start, hour_end, font_size_hour=0, draw_numbers=False)

# --- メイン関数 ---
def calendar_weekly_vertical(year, month=range(12), start_april=True, starts_with_mon=True, adjust_left=True,
                             calendar_path=None, pagesize=A5, margin=5*mm, font_path=None, font_size=12,
                             hour_start=6, hour_end=24, df_event=pd.DataFrame(), draw_day_box=False):
    if not calendar_path:
        calendar_path = f'{year}_calendar.pdf'
    df_year = create_year_df(year, start_april=start_april, starts_with_mon=starts_with_mon, adjust_left=adjust_left)
    c = canvas.Canvas(calendar_path, pagesize=pagesize)
    width, height = pagesize
    w_day = (width - 2 * margin) / 4
    h_day = height - 2 * margin
    if font_path is None:
        font_path = 'c:/Windows/Fonts/CENTURY.ttf'
    font_name = use_font(font_path)
    pdfmetrics.registerFont(TTFont(font_name, font_path))
    pages = df_year['page'].unique()
    for p in pages:
        df_page = df_year[df_year['page'] == p]
        # 1. 予定がある日の描画
        for i in range(len(df_page)):
            row = df_page.iloc[i]
            left = margin + w_day * row['position']
            top = margin + h_day
            create_day(c, left, top, w_day, h_day, font_name, font_size,
                       year=str(row['year']), month=str(row['month']), day=str(row['day']), wday=row['weekday'],
                       hour_start=hour_start, hour_end=hour_end, draw_year_month=row['draw_year_month'],
                       df_event=df_event)
        # 2. 右端の空きブロック（メモ欄）の描画
        days_per_page = 4
        existing_positions = df_page['position'].tolist()
        for pos in range(days_per_page):
            if pos not in existing_positions:
                left = margin + w_day * pos
                top = margin + h_day
                draw_empty_block(c, left, top, w_day, h_day, hour_start, hour_end)
                
        c.showPage()
    c.save()
    return calendar_path

if __name__ == '__main__':
    font_path = './HackGen35Console-Regular.ttf' # not work in font directory
    font_path = 'c:/Windows/Fonts/CENTURY.ttf'
    font_path = 'c:/Windows/Fonts/ALGER.TTF'
    font_path = 'c:/Windows/Fonts/ITCBLKAD.TTF'
    font_path = 'c:/Windows/Fonts/BRUSHSCI.TTF'
    font_path = 'c:/Windows/Fonts/UDDigiKyokashoN-R.ttc'

    year            = 2026
    hour_start      = 6
    hour_end        = 22
    starts_with_mon = True
    adjust_left     = True

    # event data
    import event
    path = 'schedule.xlsx'
    df_input = pd.read_excel(path)
    df_date = event.generate_schedule(df_input)
    df_event = event.format_events(df_date)

    path_calendar = calendar_weekly_vertical(year, 
        font_path       = font_path, 
        hour_start      = hour_start, 
        hour_end        = hour_end,
        starts_with_mon = starts_with_mon, 
        adjust_left     = adjust_left, 
        df_event        = df_event)

    os.startfile(path_calendar)
