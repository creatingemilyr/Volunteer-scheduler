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

    df.columns = df.columns.str.strip().str.lower()
    rename_map = {
        'full name': 'Full name',
        'service week available': 'Service Week Available',
        'service times available': 'Service Times Available',
        'black out dates': 'Black Out Dates'
    }
    df.rename(columns=rename_map, inplace=True)

    for col in rename_map.values():
        if col not in df.columns:
            df[col] = ""

    df_clean = df[list(rename_map.values())].dropna(subset=['Full name'])
    df_clean['Service Week Available'] = df_clean['Service Week Available'].str.lower().str.replace(' ', '')
    df_clean['Service Times Available'] = df_clean['Service Times Available'].str.lower().str.replace(' ', '').str.replace(':', '')

    blackout_dict = {}
    for _, row in df_clean.iterrows():
        name = row['Full name']
        raw_dates = str(row['Black Out Dates']).strip()
        if raw_dates and raw_dates.lower() != 'nan':
            dates = [d.strip() for d in raw_dates.split(',') if d.strip()]
            blackout_dict[name] = set(dates)
        else:
            blackout_dict[name] = set()

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
    st.success("✅ Schedule generated!")
    st.dataframe(schedule_df, use_container_width=True)

    csv = schedule_df.to_csv(index=False).encode('utf-8')
    st.download_button("📥 Download Schedule CSV", data=csv, file_name="Volunteer_Schedule.csv", mime='text/csv')

    # Monthly Summary Table
    st.markdown("### 📊 Volunteer Monthly Summary")
    summary_data = []
    months = sorted(set([d.strftime('%Y-%m') for d in all_sundays]))

    for volunteer in df_clean['Full name']:
        row = {"Volunteer": volunteer}
        for month in months:
            row[month] = volunteer_monthly_count[volunteer][month]
        summary_data.append(row)

    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True)

    # Volunteers Who Didn't Meet Minimum
    min_missed = [(row['Volunteer'], month)
                  for row in summary_data
                  for month in months if row[month] < 1]

    if min_missed:
        st.markdown("### ⚠️ Volunteers Who Didn't Meet the 1x/Month Minimum")
        missed_df = pd.DataFrame(min_missed, columns=["Volunteer", "Month"])
        st.dataframe(missed_df)
    else:
        st.success("🎉 All volunteers met the 1x/month minimum!")

    # Export to Excel with separate monthly tabs
    xls_buffer = pd.ExcelWriter("/tmp/Volunteer_Schedule_Months.xlsx", engine='openpyxl')
    schedule_df['Month'] = pd.to_datetime(schedule_df['Date']).dt.to_period('M')
    for period, group_df in schedule_df.groupby('Month'):
        month_name = period.strftime('%B %Y')
        group_df.drop(columns=['Month'], inplace=True)
        group_df.to_excel(xls_buffer, sheet_name=month_name, index=False)
    xls_buffer.close()

    with open("/tmp/Volunteer_Schedule_Months.xlsx", "rb") as f:
        st.download_button(
            label="📥 Download Excel with Monthly Tabs",
            data=f.read(),
            file_name="Volunteer_Schedule_Monthly_Tabs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
