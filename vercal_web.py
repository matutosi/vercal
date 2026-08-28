import datetime
import os

import pandas as pd
import streamlit as st

import event
import vercal

# settings
font_path = 'HackGen35Console-Regular.ttf'
if not os.path.exists(font_path):
    font_path = None

settings = st.sidebar

# ボタンは画面の上に置くが，設定を読んでから処理したいので場所だけ先に確保する
with settings:
    button_slot = st.empty()
    download_slot = st.empty()

with settings:
    # Calculate default year based on calendar start month
    now = datetime.datetime.now()

    start_april = st.checkbox('Starts with April', value=True)
    # For April-start calendar (default): Jan-Mar uses current year, Apr-Dec uses next year
    # For January-start calendar: always uses next year
    if start_april and now.month <= 3:
        default_year = now.year
    else:
        default_year = now.year + 1
    year = st.number_input('Year:', value=default_year, step=1)
    hour_start, hour_end = st.slider('Range in a day', min_value=0, max_value=24, value=(6, 24), step=1)
    starts_with_mon = st.checkbox('Starts with Monday', value=True)
    adjust_left = st.checkbox('Adjust left', value=True)
    uploaded_file = st.file_uploader('Upload a schedule file', type='xlsx')

df_event = None
if uploaded_file is not None:
    df_input = pd.read_excel(uploaded_file)
    df_date = event.generate_schedule(df_input)
    df_event = event.format_events(df_date)
    st.write(df_date)

if button_slot.button('Create Calendar'):
    if hour_end <= hour_start:
        st.error('Range in a day: 終了時刻は開始時刻より後にする')
    else:
        path_calendar = vercal.calendar_weekly_vertical(year,
            start_april     = start_april,
            font_path       = font_path,
            hour_start      = hour_start,
            hour_end        = hour_end,
            starts_with_mon = starts_with_mon,
            adjust_left     = adjust_left,
            df_event        = df_event)
        download_slot.download_button('DOWNLOAD CALENDAR', open(path_calendar, 'br'), path_calendar)
