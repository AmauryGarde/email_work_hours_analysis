# imports
import pandas as pd

# Variables
document_name = "insert_doc_name_here"  # e.g. "mail_file.xlsx"
date_column = "Date full"
subject_column = "Subject"

# Read the document
df = pd.read_excel(document_name)

# Extracting date and time
df['date'] = pd.to_datetime(df['Date full'])

df['date'] = df['Date full'].dt.date
df['time'] = df['Date full'].dt.time

# Define working hours
start_work = pd.Timestamp('08:00:00').time()
end_work = pd.Timestamp('17:00:00').time()


# Aggregation function
def calculate_hours(group):
    first_email = group.loc[group['Date full'].idxmin()]
    last_email = group.loc[group['Date full'].idxmax()]
    first_email_time = first_email['time']
    last_email_time = last_email['time']
    first_email_subject = first_email['Subject']
    last_email_subject = last_email['Subject']

    # Calculate effective start and end times within work hours
    effective_first_email_time = max(first_email_time, start_work)
    effective_last_email_time = min(last_email_time, end_work)

    # Ensure the first email time is at least 8 AM and last email time is at least 5 PM
    if effective_first_email_time > start_work:
        effective_first_email_time = start_work
    if effective_last_email_time < end_work:
        effective_last_email_time = end_work

    # Calculate total hours worked
    work_hours = (pd.Timestamp.combine(group['date'].iloc[0], effective_last_email_time) - pd.Timestamp.combine(
        group['date'].iloc[0], effective_first_email_time)).seconds / 3600
    work_hours = max(work_hours, 9)  # Ensure minimum total hours is 9 (8 AM to 5 PM)

    # Calculate extra hours worked outside the working hours
    extra_hours = 0
    if first_email_time < start_work:
        extra_hours += (pd.Timestamp.combine(group['date'].iloc[0], start_work) - pd.Timestamp.combine(
            group['date'].iloc[0], first_email_time)).seconds / 3600
    if last_email_time > end_work:
        extra_hours += (pd.Timestamp.combine(group['date'].iloc[0], last_email_time) - pd.Timestamp.combine(
            group['date'].iloc[0], end_work)).seconds / 3600

    # Adjust total hours to include extra hours
    total_hours = work_hours + extra_hours

    return pd.Series({
        'first_email_hour': first_email_time,
        'first_email_subject': first_email_subject,
        'last_email_hour': last_email_time,
        'last_email_subject': last_email_subject,
        'total_hours_worked': total_hours,
        'extra_hours_worked': extra_hours
    })


# Apply aggregation
summary = df.groupby('date').apply(calculate_hours).reset_index()

# Adding totals row
totals = pd.DataFrame({
    'date': ['Totals'],
    'first_email_hour': [''],
    'first_email_subject': [''],
    'last_email_hour': [''],
    'last_email_subject': [''],
    'total_hours_worked': [summary['total_hours_worked'].sum()],
    'extra_hours_worked': [summary['extra_hours_worked'].sum()]
})

summary = pd.concat([summary, totals], ignore_index=True)

# Export to Excel
summary.to_excel('email_analysis.xlsx', index=False)
