import streamlit as st
import pandas as pd
from datetime import date
from collections import defaultdict

st.set_page_config(page_title="Volunteer Scheduler", layout="wide")
st.title("Church Volunteer Scheduler")

uploaded_file = st.file_uploader("Upload volunteer availability file (.csv or .xlsx)", type=["csv", "xlsx"])
range_option = st.selectbox("Schedule for how long?", ["1 month", "2 months", "3 months"])
range_months = int(range_option.split()[0])
start_date = st.date_input("Select the start date for the schedule", value=date.today())

if uploaded_file and st.button("Generate Schedule"):
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Normalize columns: remove spaces, make lowercase
    df.columns = df.columns.str.strip().str.lower()

    # Rename to expected casing
    rename_map = {
        'full name': 'Full name',
        'service week available': 'Service Week Available',
        'service times available': 'Service Times Available',
        'black out dates': 'Black Out Dates'
    }
    df.rename(columns=rename_map, inplace=True)

    # Add missing columns if needed
    for col in rename_map.values():
        if col not in df.columns:
            df[col] = ""

    # Clean and prepare
    df_clean = df[list(rename_map.values())].dropna(subset=['Full name'])
    df_clean['Service Week Available'] = df_clean['Service Week Available'].str.lower().str.replace(' ', '')
    df_clean['Service Times Available'] = df_clean['Service Times Available'].str.lower().str.replace(' ', '').str.replace(':', '')

    # Build blackout date dictionary
    blackout_dict = {}
    for _, row in df_clean.iterrows():
        name = row['Full name']
        raw_dates = str(row['Black Out Dates']).strip()
        if raw_dates and raw_dates.lower() != 'nan':
            dates = [d.strip() for d in raw_dates.split(',') if d.strip()]
            blackout_dict[name] = set(dates)
        else:
            blackout_dict[name] = set()

    # Schedule range
    end_date = start_date + pd.DateOffset(months=range_months)
    all_sundays = pd.date_range(start=start_date, end=end_date, freq='W-SUN')

    week_labels = ['1stsunday', '2ndsunday', '3rdsunday', '4thsunday', '5thsunday']
    week_number_map = {i: week_labels[i % 5] for i in range(len(all_sundays))}

    volunteer_monthly_count = defaultdict(lambda: defaultdict(int))
    schedule_data = []

    for i, sunday in enumerate(all_sundays):
        week_label = week_number_map[i]
        month = sunday.strftime('%Y-%m')
        date_str = sunday.strftime('%Y-%m-%d')
        for service_time in ['8am', '930am', '11am']:
            eligible = df_clean[
                df_clean['Service Week Available'].str.contains(week_label) &
                df_clean['Service Times Available'].str.contains(service_time)
            ]['Full name'].tolist()

            selected = []
            for volunteer in eligible:
                if volunteer_monthly_count[volunteer][month] < 2 and date_str not in blackout_dict.get(volunteer, set()):
                    selected.append(volunteer)
                if len(selected) == 2:
                    break

            for slot in range(2):
                if selected:
                    volunteer = selected.pop(0)
                    volunteer_monthly_count[volunteer][month] += 1
                else:
                    volunteer = "Volunteer Needed"
                schedule_data.append({
                    'Date': date_str,
                    'Week': week_label,
                    'Service Time': service_time.upper().replace('AM', ' AM'),
                    'Volunteer Slot': f"Slot {slot+1}",
                    'Volunteer': volunteer
                })

    schedule_df = pd.DataFrame(schedule_data)
    st.success("Schedule generated!")
    st.dataframe(schedule_df, use_container_width=True)

    csv = schedule_df.to_csv(index=False).encode('utf-8')
    st.download_button("Download Schedule CSV", data=csv, file_name="Volunteer_Schedule.csv", mime='text/csv')
