import os
import calendar
from datetime import datetime, timedelta
import pandas as pd

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5
from reportlab.lib.units import mm

def create_year_df(year,
    start_april=True, starts_with_mon=True, adjust_left=True):
    df_year = create_year(year, start_april=start_april)
    df_year = add_position(df_year, starts_with_mon=starts_with_mon, adjust_left=adjust_left)
    df_year = add_page(df_year)
    df_year = add_draw_year_month(df_year)
    return df_year

def create_year(year, start_april=True):
    # start with January
    if not start_april:
        return generate_dates(year)
    # start with April
    cal_this_yr = generate_dates(year).query('month > 3')
    cal_next_yr = generate_dates(year + 1).query('month < 4')
    calender = pd.concat([cal_this_yr, cal_next_yr])
    return calender

def generate_dates(year):
    """
    Generates a Pandas DataFrame with year, month, day, and weekday.
    Examples
        year_to_generate = 2025
        date_df = generate_dates(year_to_generate)
        print(date_df)
    """
    data = []
    for month in range(1, 13):
        for day in range(1, calendar.monthrange(year, month)[1] + 1):
            weekday_num = calendar.weekday(year, month, day)
            weekday_abbr = calendar.day_abbr[weekday_num].lower()
            data.append([year, month, day, weekday_abbr])
    df = pd.DataFrame(data, columns=['year', 'month', 'day', 'weekday'])
    return df

def add_page(df, base_col='position', page_col='page'):
    df[page_col] = 1
    for i in range(1, len(df)):
        if df[base_col].iloc[i] < df[base_col].iloc[i - 1]:
            df.loc[i:, page_col] += 1
    return df

def add_position(df, starts_with_mon=True, adjust_left=True):
    if starts_with_mon:
        if adjust_left:
            data = {'weekday': ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'], 
                    'position': [0, 1, 2, 3, 0, 1, 2]}
        else:
            data = {'weekday': ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'], 
                    'position': [1, 2, 3, 0, 1, 2, 3]}
    else:
        if adjust_left:
            data = {'weekday': ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'], 
                    'position': [0, 1, 2, 3, 0, 1, 2]}
        else:
            data = {'weekday': ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'], 
                    'position': [1, 2, 3, 0, 1, 2, 3]}
    wday = pd.DataFrame(data)
    return df.merge(wday, how='left')

def add_draw_year_month(df):
    df['draw_year_month'] = False
    df.loc[df['day'] == 1, 'draw_year_month'] = True
    for i in range(1, len(df)):
        if df['position'].iloc[i] < df['position'].iloc[i - 1]:
            df.loc[i, 'draw_year_month'] = True
    return df

def create_day(c, 
               left, top, width, height, 
               font_name, font_size,
               year, month, day, wday,
               hour_start=6, hour_end=24, 
               h_year_month=5*mm, h_wday_day=5*mm, h_memo=15*mm,
               draw_year_month=True,
               df_event=pd.DataFrame(),
               draw_day_box=False):
    """Creates a daily schedule block."""
    offset = 1 * mm
    leftt = left + offset
    right = left + width - offset
    bottom = top - height
    c.setFont(font_name, font_size)
    font_size_hour = font_size * 0.7
    top_hour = top - (h_year_month + h_wday_day + h_memo)
    h_hour = top_hour - bottom
    # settings = {"left"           : left         , 
    #             "right"          : right        ,
    #             "top"            : top          ,
    #             "top_hour"       : top_hour     ,
    #             "year"           : year         ,
    #             "month"          : month        ,
    #             "day"            : day          ,
    #             "wday"           : wday         ,
    #             "h_year_month"   : h_year_month ,
    #             "h_wday_day"     : h_wday_day   ,
    #             "h_hour"         : h_hour       ,
    #             "h_memo"         : h_memo       ,
    #             "hour_start"     : hour_start   ,
    #             "hour_end"       : hour_end     ,
    #             "font_size_hour": font_size_hour,
    #             }
    date_section(c, left,        top,      year, month, day, wday, h_year_month, h_wday_day, draw_year_month)
    memo_section(c, left, right, top,                              h_year_month, h_wday_day, h_memo)
    ten_minute  (c, left, right, top_hour, h_hour, hour_start, hour_end)
    hour_section(c, left, right, top_hour, h_hour, hour_start, hour_end, font_size_hour)
    date_today = datetime.strptime(f'{year}-{month}-{day}', '%Y-%m-%d').strftime('%Y-%m-%d')
    print(f'drawing {date_today}')
    events = df_event['event'][df_event['date'] == date_today]
    if not len(events) == 0:
        for event in events.iloc[0]:
            schedule    = event['event']
            event_start = event['event_start']
            event_end   = event['event_end']
            draw_schedule(c, schedule, event_start, event_end, 
                hour_start, hour_end, 
                top_hour, left, 
                width, h_hour, 
                font_size_hour)
    if draw_day_box:
        day_box(c, left, top, width, height)

def draw_schedule(c, schedule, event_start, event_end, hour_start, hour_end, top_hour, left, width, h_hour, font_size_hour):
    hours = hour_end - hour_start
    one_hour = h_hour / hours
    x = left
    y = top_hour - (string2float(event_start) - hour_start) * one_hour
    duration = - (string2float(event_end) - string2float(event_start)) * one_hour
    c.setDash([1, 0]) # [on, off]
    c.setLineWidth(0.3)
    c.rect(x + width * 0.12, y, width * 0.83, duration)
    c.setFont(c._fontname, font_size_hour)
    c.drawString(x + width * 0.13, y - font_size_hour, schedule)

def hour_section(c, left, right, top_hour, h_hour, hour_start, hour_end, font_size_hour):
    """Draws hour lines and times."""
    c.setDash([3, 2]) # [on, off]
    c.setLineWidth(0.8)
    c.setFont(c._fontname, font_size_hour)
    hours = hour_end - hour_start
    one_hour = h_hour / hours
    for i in range(hours):
        c.drawString(left, top_hour - one_hour * i - font_size_hour, str(hour_start + i).zfill(2))
    for i in range(1, hours + 1):
        c.line(left, top_hour - one_hour * i, right, top_hour - one_hour * i)

def string2float(time_str):
    """convert time string into time float """
    hours, minutes = map(int, time_str.split(':'))
    return hours + minutes / 60.0

def use_font(font_path):
    file_name = os.path.basename(font_path)
    font_name, _ = os.path.splitext(file_name)
    return font_name

def date_section(c, left, top, year, month, day, wday, h_year_month, h_wday_day, draw_year_month=True):
    """Draws date information (year-month, day day of the week)."""
    if draw_year_month:
        c.drawString(left, top - c._fontsize, f'{year}-{month.zfill(2)}')
    c.drawString(left, top - h_wday_day - c._fontsize, f'{day.zfill(2)} {wday}')

def memo_section(c, left, right, top, h_year_month, h_wday_day, h_memo, circle_no=3):
    """Draws memo section (separator line and circles)."""
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
    """Draws circles every 10 minutes."""
    hours = hour_end - hour_start
    one_hour = h_hour / hours
    ten_minutes = one_hour / 6
    x = left - (left - right) * 0.12
    c.setDash([1, 0])
    c.setLineWidth(0.8)
    for i in range((hours) * 6):
        y = top - ten_minutes * i
        circle_size = ((i % 2) + 1) * 0.5 # even: large, odd: small
        c.circle(x, y, circle_size)

def day_box(c, left, top, width, height):
    """Draws a frame for debug."""
    c.setLineWidth(1)
    c.setDash([1, 0]) # [on, off]
    c.line(left, top, left + width, top)
    c.line(left, top - height, left + width, top - height)
    c.line(left, top, left, top - height)
    c.line(left + width, top, left + width, top - height)

def calendar_weekly_vertical(year, month=range(12),
    start_april=True, starts_with_mon=True, adjust_left=True,
    calendar_path=None, 
    pagesize=A5, margin=5*mm,
    font_path=None, font_size=12,
    hour_start=6, hour_end=24, 
    df_event = pd.DataFrame(),
    draw_day_box=False):
    # # # # settings # # # #
    # output path
    if not calendar_path:
        calendar_path = f'{year}_calendar.pdf'
    df_year = create_year_df(year, start_april=start_april, starts_with_mon=starts_with_mon, adjust_left=adjust_left)
    # page size
    c = canvas.Canvas(calendar_path, pagesize=pagesize)
    width, height = pagesize
    w_day = (width - 2 * margin) / 4
    h_day = height - 2 * margin
    # font
    if font_path is None:
        font_path = 'c:/Windows/Fonts/CENTURY.ttf'
    font_name = use_font(font_path)
    pdfmetrics.registerFont(TTFont(font_name, font_path))
    # # # # create pages # # # #
    pages = df_year['page'].unique()
    for p in pages:
        df_page = df_year[df_year['page'] == p]
        for i in range(len(df_page)):
            year, month, day, weekday, position, page, draw_year_month = df_page.iloc[i]
            left = margin + w_day * position
            top  = margin + h_day
            create_day(c, 
                       left, top, w_day, h_day,
                       font_name, font_size,
                       year=str(year), month=str(month), day=str(day), wday=weekday,
                       hour_start=hour_start, hour_end=hour_end,
                       h_year_month=5*mm, h_wday_day=5*mm, h_memo=15*mm,
                       draw_year_month=draw_year_month,
                       df_event=df_event,
                       draw_day_box=draw_day_box)
        c.showPage()
    c.save()
    return calendar_path

# 
if __name__ == '__main__':

    font_path = 'D:/matu/work/ToDo/kwu/calendar/GenShinGothic-Monospace-Medium.ttf'
    font_path = 'c:/Windows/Fonts/CENTURY.ttf'
    font_path = 'c:/Windows/Fonts/ALGER.TTF'
    font_path = 'c:/Windows/Fonts/ITCBLKAD.TTF'
    font_path = 'c:/Windows/Fonts/BRUSHSCI.TTF'
    font_path = 'c:/Windows/Fonts/UDDigiKyokashoN-R.ttc'
    font_path = 'D:/matu/work/ToDo/vercal/HackGen35Console-Regular.ttf' # not work in font directory

    year            = 2025
    hour_start      = 6
    hour_end        = 22
    starts_with_mon = True
    adjust_left     = True

    # event data
    import event
    path = 'event.xlsx'
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
