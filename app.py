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

    # Normalize and rename columns
    df.columns = df.columns.str.strip().str.lower()
    rename_map = {
        'full name': 'Full name',
        'service week available': 'Service Week Available',
        'service times available': 'Service Times Available',
        'service times avaliable': 'Service Times Available',  # common typo fix
        'black out dates': 'Black Out Dates'
    }
    df.rename(columns=rename_map, inplace=True)

    expected_cols = ['Full name', 'Service Week Available', 'Service Times Available', 'Black Out Dates']
    for col in expected_cols:
        if col not in df.columns:
            df[col] = ""

    df_clean = df[expected_cols].dropna(subset=['Full name'])
    df_clean['Service Week Available'] = df_clean['Service Week Available'].str.lower().str.replace(" ", "").str.replace(",", ",")
    df_clean['Service Times Available'] = df_clean['Service Times Available'].str.lower().str.replace(" ", "").str.replace(":", "").str.replace(",", ",")

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
    volunteer_monthly_count = defaultdict(lambda: defaultdict(int))
    schedule_data = []

    for sunday in all_sundays:
        week_of_month = ((sunday.day - 1) // 7) + 1
        week_label = week_labels[week_of_month - 1] if week_of_month <= 5 else '5thsunday'
        month = sunday.strftime('%Y-%m')
        date_str = sunday.strftime('%Y-%m-%d')
        for service_time in ['8am', '930am', '11am']:
            eligible = []
            for _, row in df_clean.iterrows():
                if week_label in row['Service Week Available'] and service_time in row['Service Times Available']:
                    if date_str not in blackout_dict.get(row['Full name'], set()):
                        if volunteer_monthly_count[row['Full name']][month] < 2:
                            eligible.append(row['Full name'])

            selected = eligible[:2]

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

    min_missed = [(row['Volunteer'], month)
                  for row in summary_data
                  for month in months if row[month] < 1]

    if min_missed:
        st.markdown("### ⚠️ Volunteers Who Didn't Meet the 1x/Month Minimum")
        missed_df = pd.DataFrame(min_missed, columns=["Volunteer", "Month"])
        st.dataframe(missed_df)
    else:
        st.success("🎉 All volunteers met the 1x/month minimum!")

    # Excel export with monthly tabs
    xls_path = "/tmp/Volunteer_Schedule_Monthly_Tabs.xlsx"
    with pd.ExcelWriter(xls_path, engine='openpyxl') as writer:
        schedule_df['Month'] = pd.to_datetime(schedule_df['Date']).dt.to_period('M')
        for period, group_df in schedule_df.groupby('Month'):
            month_name = period.strftime('%B %Y')
            group_df.drop(columns=['Month'], inplace=True)
            group_df.to_excel(writer, sheet_name=month_name, index=False)

    with open(xls_path, "rb") as f:
        st.download_button(
            label="📥 Download Excel with Monthly Tabs",
            data=f.read(),
            file_name="Volunteer_Schedule_Monthly_Tabs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
