import os
import datetime

import pandas as pd
import streamlit as st

import event
import vercal

# setgings
hour_end  = 24
df_event  = pd.DataFrame(columns=["date", "week_of_day","event_start","event_end","event"])
font_path = 'HackGen35Console-Regular.ttf'

settings = st.sidebar

def create_calender():
    path_calendar = vercal.calendar_weekly_vertical(year, 
        start_april     = start_april,
        font_path       = font_path, 
        hour_start      = hour_start, 
        hour_end        = hour_end,
        starts_with_mon = starts_with_mon, 
        adjust_left     = adjust_left, 
        df_event        = df_event)
    st.download_button('DOWNLOAD CALENDAR', open(path_calendar, 'br'), path_calendar)

with settings:
    st.button('Create Calendar', on_click=create_calender)
    year                 = st.number_input('Year:', value=datetime.datetime.now().year, step=1)
    start_april          = st.checkbox('Starts with April', value=True)
    hour_start, hour_end = st.slider("Range in a day", min_value=0, max_value=24, value=(6, 24), step=1)
    starts_with_mon      = st.checkbox('Starts with Monday', value=True)
    adjust_left          = st.checkbox('Adjust left'       , value=True)
    uploaded_file = st.file_uploader("Upload a schedule file", type="xlsx")

if uploaded_file is not None:
    df_input = pd.read_excel(uploaded_file)
    df_date  = event.generate_schedule(df_input)
    df_event = event.format_events(df_date)
    st.write(df_date)
