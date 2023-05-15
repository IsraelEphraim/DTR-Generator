import os
import csv
from flask import Flask, render_template, request, redirect, send_file, jsonify, url_for
import pandas as pd
from datetime import datetime, timedelta, time

port = int(os.environ.get('PORT', 9000))

app = Flask(__name__, template_folder='template')

def delete_duplicate(df):
    df.drop_duplicates(subset=['Employee Code', 'Date'], inplace=True)
    df.to_csv('dtr.csv', index=False)
    return df

@app.route('/')
def index():
    csv_file_path = 'dtr.csv'

    if os.path.exists(csv_file_path):
        os.remove(csv_file_path)

    try:
        dtr = pd.read_csv('dtr.csv')
    except FileNotFoundError:
        dtr = pd.DataFrame(columns=['Status',
                                    'Overtime',
                                    'Employee Code',
                                    'Employee Name',
                                    'Date',
                                    'Day',
                                    'Weekday or Weekend',
                                    'Work Description',
                                    'Time In',
                                    'Time Out',
                                    'Actual Time In',
                                    'Actual Time Out',
                                    'Net Hours Rendered (Time Format)',
                                    'Actual Gross Hours Render',
                                    'Hours Rendered',
                                    'Undertime Hours',
                                    'Tardiness',
                                    'Excess of 8 hours Overtime',
                                    'Total of 8 hours Overtime',
                                    'Night Difference',
                                    'Night Difference Overtime',
                                    'Night Difference First 8 hours',
                                    'Night Difference Excess of 8 hours'
                                    ])

    # Read the DataFrame
    try:
        df = pd.read_csv('dtr.csv')
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Status',
                                    'Overtime',
                                    'Employee Code',
                                    'Employee Name',
                                    'Date',
                                    'Day',
                                    'Weekday or Weekend',
                                    'Work Description',
                                    'Time In',
                                    'Time Out',
                                    'Actual Time In',
                                    'Actual Time Out',
                                    'Net Hours Rendered (Time Format)',
                                    'Actual Gross Hours Render',
                                    'Hours Rendered',
                                    'Undertime Hours',
                                    'Tardiness',
                                   'Excess of 8 hours Overtime',
                                   'Total of 8 hours Overtime',
                                    'Night Difference',
                                    'Night Difference Overtime',
                                   'Night Difference First 8 hours',
                                   'Night Difference Excess of 8 hours'])
    delete_duplicate(df)

    return render_template('index.html', dtr=dtr)

#Auto Generate of Name using Employee Code
@app.route("/lookup_employee", methods=["POST"])
def lookup_employee():
    employee_data = pd.read_csv("dtr.csv")
    employee_code = request.json["employee_code"]
    employee_name = employee_data.loc[employee_data["Employee Code"] == employee_code, "Employee Name"].iloc[0]
    return jsonify({"employee_name": employee_name})

def calculate_timeanddate(employee_name, employee_code, day_of_week, date_transact1, date_obj, work_descript, time_in, time_out, time_in1, time_out1, actual_time_in, actual_time_out, actual_time_in1, actual_time_out1):
    week_day = date_obj.weekday()
    if week_day < 5:  # Weekday (Monday to Friday)
        week_check = "Weekday"
    else:  # Weekend day (Saturday or Sunday)
        week_check = "Weekend"
        # Do something with day_of_week variable



    datetime_timein = datetime.combine(date_obj.date(), time_in1)
    if time_in1 > time_out1:
        add_day = timedelta(days=1)
        new_day = date_obj.date() + add_day
        datetime_timeout = datetime.combine(new_day, time_out1)

    else:
        datetime_timeout = datetime.combine(date_obj.date(), time_out1)

    total_datetime_in_out = datetime_timeout - datetime_timein

    actual_datetime_timein = datetime.combine(date_obj.date(), actual_time_in1)
    if actual_time_in1 > actual_time_out1:
        actual_add_day = timedelta(days=1)
        actual_new_day = date_obj.date() + actual_add_day
        actual_datetime_timeout = datetime.combine(actual_new_day, actual_time_out1)

    else:
        actual_datetime_timeout = datetime.combine(date_obj.date(), actual_time_out1)

    total_actual_datetime_in_out =  actual_datetime_timeout - actual_datetime_timein

    # Non Chargeable Break
    non_chargeable_break = time(1, 0, 0)
    non_chargeable_break_timedelta = timedelta(hours=non_chargeable_break.hour, minutes=non_chargeable_break.minute,
                                               seconds=non_chargeable_break.second)

    # Shift Hours
    shift_hour = datetime_timeout - datetime_timein
    shift_hour_str = str(shift_hour)

    # Net Hours Rendered
    net_hours_rendered = shift_hour - non_chargeable_break_timedelta
    net_hours_rendered_str = str(net_hours_rendered)
    net_hours_rendered_time = round(net_hours_rendered / timedelta(hours=1), 2)


    # Actual Gross Hours Rendered
    if total_actual_datetime_in_out > non_chargeable_break_timedelta:
        actual_render = total_actual_datetime_in_out - non_chargeable_break_timedelta
        actual_render_str = str(actual_render)
        actual_render_str = datetime.strptime(actual_render_str, '%H:%M:%S')
        actual_render_str = actual_render_str.strftime('%H.%M')
        actual_render2 = float(actual_render_str)
    else:
        actual_render = total_actual_datetime_in_out
        actual_render_str = str(actual_render)
        actual_render_str = datetime.strptime(actual_render_str, '%H:%M:%S')
        actual_render_str = actual_render_str.strftime('%H.%M')
        actual_render2 = float(actual_render_str)

    # Regular 8 Hours
    regular = time(8, 0, 0)
    regular_timedelta = timedelta(hours=regular.hour, minutes=regular.minute,
                                  seconds=regular.second)
    if actual_render > regular_timedelta:
        regular_8 = regular_timedelta
        regular_8 = str(regular_8)

    else:
        regular_8 = timedelta(hours=round(actual_render.total_seconds() / 3600))
        regular_8 = str(regular_8)


    # Hours Rendered
    hour_rendered = total_actual_datetime_in_out
    hour_rendered_str = str(hour_rendered)
    hour_rendered_str = datetime.strptime(hour_rendered_str, '%H:%M:%S')
    hour_rendered_str = hour_rendered_str.strftime('%H.%M')
    hour_rendered_str = float(hour_rendered_str)


    # Total Hours
    total_hours = hour_rendered - non_chargeable_break_timedelta
    total_hours_str = round(total_hours / timedelta(hours=1), 2)
    if total_hours_str < 0:
        total_hours_str = 0

    # Overtime
    overtime = total_hours - net_hours_rendered
    overtime_str = round(overtime / timedelta(hours=1), 2)

    if overtime_str < 1:
        overtime_str = 0

    excess_overtime = 0
    total_8hours = 0
    if overtime_str > 8:
        excess_overtime = overtime_str - 8
        total_8hours = overtime_str - excess_overtime

    # Excess Hours in Numerical
    excess_num = total_hours - net_hours_rendered
    excess_num = round(excess_num / timedelta(hours=1), 2)
    if excess_num < 0:
        excess_num = 0


    # Undertime Hours
    if total_actual_datetime_in_out < total_datetime_in_out:
        undertime_check_try = total_datetime_in_out - total_actual_datetime_in_out
        undertime_check_try = str(undertime_check_try)
        undertime_check_try = datetime.strptime(undertime_check_try, '%H:%M:%S')
        undertime_check_try = undertime_check_try.strftime('%H.%M')
        undertime_check_try = float(undertime_check_try)
    else:
        undertime_check_try = 0

    # Time in Tardiness
    expected = time(0, 0, 0)
    expected_timedelta = timedelta(hours=expected.hour, minutes=expected.minute, seconds=expected.second)

    if actual_datetime_timein > datetime_timein:
        tardiness = actual_datetime_timein - datetime_timein
        tardiness_str = str(tardiness)
        tardiness_str = datetime.strptime(tardiness_str, '%H:%M:%S')
        tardiness_str = tardiness_str.strftime('%H.%M')
        tardiness_str = float(tardiness_str)

    else:
        tardiness = expected_timedelta
        tardiness_str = str(tardiness)
        tardiness_str = datetime.strptime(tardiness_str, '%H:%M:%S')
        tardiness_str = tardiness_str.strftime('%H.%M')
        tardiness_str = float(tardiness_str)

    # Night Difference
    night_diff_start = '22:00'
    night_diff_end = '06:00'

    datetime_start_night = datetime.strptime(night_diff_start, "%H:%M").time()
    datetime_end_night = datetime.strptime(night_diff_end, "%H:%M").time()
    date_time_start_night = datetime.combine(date_obj, datetime_start_night)
    datetime_end_before_night = datetime.combine(date_obj, datetime_end_night)
    one_day = timedelta(days=1)
    new_date = date_obj + one_day
    date_time_end_night = datetime.combine(new_date, datetime_end_night)

    date_time_actual_start = datetime.combine(date_obj, actual_time_in1)
    datetime_actual_end = datetime.combine(date_obj, actual_time_out1)

    night_date = date_obj
    if date_time_actual_start > datetime_actual_end:
        night_date = date_obj + one_day
    date_time_actual_end = datetime.combine(night_date, actual_time_out1)

    night_diff_total = 0
    if date_time_actual_start >= date_time_start_night and date_time_actual_end <= date_time_end_night and date_time_actual_end > date_time_start_night:
        night_diff_total = date_time_actual_end - date_time_actual_start
        night_diff_total = str(night_diff_total)
        night_diff_total = datetime.strptime(night_diff_total, '%H:%M:%S')
        night_diff_total = night_diff_total.strftime('%H.%M')
        night_diff_total = float(night_diff_total)
    elif date_time_actual_start > date_time_start_night and date_time_actual_end > date_time_end_night:
        night_diff_total = date_time_actual_start - date_time_end_night
        night_diff_total = str(night_diff_total)
        night_diff_total = datetime.strptime(night_diff_total, '%H:%M:%S')
        night_diff_total = night_diff_total.strftime('%H.%M')
        night_diff_total = float(night_diff_total)
    elif date_time_actual_start < date_time_start_night and date_time_actual_end <= date_time_end_night and date_time_actual_end > date_time_start_night:
        night_diff_total = date_time_actual_end - date_time_start_night
        night_diff_total = str(night_diff_total)
        night_diff_total = datetime.strptime(night_diff_total, '%H:%M:%S')
        night_diff_total = night_diff_total.strftime('%H.%M')
        night_diff_total = float(night_diff_total)
    elif date_time_actual_start < datetime_end_before_night and date_time_actual_end > datetime_end_before_night:
        night_diff_total = datetime_end_before_night - date_time_actual_start
        night_diff_total = str(night_diff_total)
        night_diff_total = datetime.strptime(night_diff_total, '%H:%M:%S')
        night_diff_total = night_diff_total.strftime('%H.%M')
        night_diff_total = float(night_diff_total)
    elif date_time_actual_start < datetime_end_before_night and date_time_actual_end <= datetime_end_before_night:
        night_diff_total = date_time_actual_end - date_time_actual_start
        night_diff_total = str(night_diff_total)
        night_diff_total = datetime.strptime(night_diff_total, '%H:%M:%S')
        night_diff_total = night_diff_total.strftime('%H.%M')
        night_diff_total = float(night_diff_total)
    elif date_time_actual_start <= date_time_start_night and date_time_actual_end > date_time_end_night:
        night_diff_total = date_time_end_night - date_time_start_night
        night_diff_total = str(night_diff_total)
        night_diff_total = datetime.strptime(night_diff_total, '%H:%M:%S')
        night_diff_total = night_diff_total.strftime('%H.%M')
        night_diff_total = float(night_diff_total)

    # Night Difference END

    night_diff_total_overtime = 0
    hour_night_rendered = actual_render2
    night_in = datetime.combine(date_obj, time_in1)
    night_out = datetime.combine(date_obj, time_out1)
    night_date_in_out = date_obj
    if night_in > night_out:
        night_date_in_out = date_obj + one_day
    night_out = datetime.combine(night_date_in_out, time_out1)

    # Night Difference OT
    if hour_night_rendered > 8 and date_time_actual_start < date_time_start_night and date_time_actual_end > date_time_end_night and night_out < date_time_start_night:
        excess_night = date_time_end_night - date_time_start_night
        excess_night = str(excess_night)
        excess_night = datetime.strptime(excess_night, '%H:%M:%S')
        excess_night = excess_night.strftime('%H.%M')
        excess_night = float(excess_night)
        night_diff_total_overtime = excess_night

    elif hour_night_rendered > 8 and date_time_actual_end >= date_time_start_night and date_time_actual_end > night_out and night_out < date_time_start_night:
        night_diff_overtime = date_time_actual_end - date_time_start_night
        night_diff_overtime = str(night_diff_overtime)
        night_diff_overtime = datetime.strptime(night_diff_overtime, '%H:%M:%S')
        night_diff_overtime = night_diff_overtime.strftime('%H.%M')
        night_diff_overtime = float(night_diff_overtime)
        night_diff_total_overtime = night_diff_overtime

    elif hour_night_rendered > 8 and date_time_actual_end < date_time_end_night and date_time_actual_end > night_out and night_out > date_time_start_night:
        excess_night = date_time_actual_end - night_out
        excess_night = str(excess_night)
        excess_night = datetime.strptime(excess_night, '%H:%M:%S')
        excess_night = excess_night.strftime('%H.%M')
        excess_night = float(excess_night)
        night_diff_total_overtime = excess_night

    elif hour_night_rendered > 8 and date_time_actual_start < datetime_end_before_night and date_time_actual_start <= night_in and night_in > datetime_end_before_night:
        night_diff_overtime = datetime_end_before_night - date_time_actual_start
        night_diff_overtime = str(night_diff_overtime)
        night_diff_overtime = datetime.strptime(night_diff_overtime, '%H:%M:%S')
        night_diff_overtime = night_diff_overtime.strftime('%H.%M')
        night_diff_overtime = float(night_diff_overtime)
        night_diff_total_overtime = night_diff_overtime

    elif hour_night_rendered > 8 and date_time_actual_start < datetime_end_before_night and date_time_actual_start <= night_in and night_in <= datetime_end_before_night:
        excess_night = night_in - date_time_actual_start
        excess_night = str(excess_night)
        excess_night = datetime.strptime(excess_night, '%H:%M:%S')
        excess_night = excess_night.strftime('%H.%M')
        excess_night = float(excess_night)
        night_diff_total_overtime = excess_night

    #Excess of 8 hours in Night Difference
    if total_actual_datetime_in_out > total_datetime_in_out:
        excess_hours_night = total_actual_datetime_in_out - total_datetime_in_out
        excess_hours_night = str(excess_hours_night)
        excess_hours_night = datetime.strptime(excess_hours_night, '%H:%M:%S')
        excess_hours_night = excess_hours_night.strftime('%H.%M')
        excess_hours_night = float(excess_hours_night)
        excess_hours_night = excess_hours_night - 1

    else:
        excess_hours_night = 0

    night_diff_first_8 = 0
    night_diff_8 = 0
    if total_datetime_in_out < total_actual_datetime_in_out:
        if excess_hours_night > 8 and date_time_actual_end > date_time_start_night:
            night_diff_8 = date_time_actual_end - date_time_start_night
            night_diff_8 = str(night_diff_8)
            night_diff_8 = datetime.strptime(night_diff_8, '%H:%M:%S')
            night_diff_8 = night_diff_8.strftime('%H.%M')
            night_diff_8 = float(night_diff_8)
            night_diff_8 = night_diff_8 - 8
            night_diff_first_8 = night_diff_total - night_diff_8


    # Approval
    status_OT = ''
    if overtime_str >= 1:
        status_OT = 'Not Approved'
    elif overtime_str == 0:
        status_OT = 'No OT'

    # Encoding of inputs to the dataframe
    dtr_new = pd.DataFrame({'Status': [status_OT],
                            'Overtime': [overtime_str],
                            'Employee Code': [employee_code],
                            'Employee Name': [employee_name],
                            'Date': [date_transact1],
                            'Day': [day_of_week],
                            'Weekday or Weekend': [week_check],
                            'Work Description': [work_descript],
                            'Time In': [time_in],
                            'Time Out': [time_out],
                            'Actual Time In': [actual_time_in],
                            'Actual Time Out': [actual_time_out],
                            'Net Hours Rendered (Time Format)': [net_hours_rendered_str],
                            'Actual Gross Hours Render': [actual_render_str],
                            'Hours Rendered': [hour_rendered_str],
                            'Undertime Hours': [undertime_check_try],
                            'Tardiness': [tardiness_str],
                            'Excess of 8 hours Overtime': [excess_overtime],
                            'Total of 8 hours Overtime': [total_8hours],
                            'Night Difference': [night_diff_total],
                            'Night Difference Overtime': [night_diff_total_overtime],
                            'Night Difference First 8 hours': [night_diff_first_8],
                            'Night Difference Excess of 8 hours': [night_diff_8]})

    try:
        dtr = pd.read_csv('dtr.csv')
    except FileNotFoundError:
        dtr = pd.DataFrame(columns=['Status',
                                    'Overtime',
                                    'Employee Code',
                                    'Employee Name',
                                    'Date',
                                    'Day',
                                    'Weekday or Weekend',
                                    'Work Description',
                                    'Time In',
                                    'Time Out',
                                    'Actual Time In',
                                    'Actual Time Out',
                                    'Net Hours Rendered (Time Format)',
                                    'Actual Gross Hours Render',
                                    'Hours Rendered',
                                    'Undertime Hours',
                                    'Tardiness',
                                    'Excess of 8 hours Overtime',
                                    'Total of 8 hours Overtime',
                                    'Night Difference',
                                    'Night Difference Overtime',
                                    'Night Difference First 8 hours',
                                    'Night Difference Excess of 8 hours'])

    dtr = pd.concat([dtr, dtr_new], ignore_index=True)
    dtr.to_csv('dtr.csv', index=False)

    return dtr

@app.route('/upload', methods=['POST'])
def upload():
    # Get the uploaded file
    uploaded_file = request.files['file']
    # Read the Excel file into a pandas dataframe
    df = pd.read_excel(uploaded_file, header=[0], engine='openpyxl')

    dtr = pd.DataFrame(columns=['Status',
                                'Overtime',
                                'Employee Code',
                                'Employee Name',
                                'Date',
                                'Day',
                                'Weekday or Weekend',
                                'Work Description',
                                'Time In',
                                'Time Out',
                                'Actual Time In',
                                'Actual Time Out',
                                'Net Hours Rendered (Time Format)',
                                'Actual Gross Hours Render',
                                'Hours Rendered',
                                'Undertime Hours',
                                'Tardiness',
                                'Excess of 8 hours Overtime',
                                'Total of 8 hours Overtime',
                                'Night Difference',
                                'Night Difference Overtime',
                                'Night Difference First 8 hours',
                                'Night Difference Excess of 8 hours'])

    for i, row in df.iterrows():
        # Employee inputs
        employee_name = row['Employee Name']
        employee_code = row['Employee Code']


        # Date of transaction
        date_transact = str(row['Date'])
        day_of_week = ""
        date_transact1 = ""
        if date_transact is not None:
            date_obj = datetime.strptime(date_transact, "%Y-%m-%d %H:%M:%S")
            date_transact1 = date_obj.date()
            day_of_week = date_obj.strftime("%A")
        else:
            date_obj = ""
            # Handle the case where the date is missing or empty

        # Description inputs
        work_descript = row['Work Description'].lower()
        if work_descript.lower() == 'regular day':
            work_descript = 'regular day'
        elif work_descript.lower() == 'legal holiday':
            work_descript = 'legal holiday'
        elif work_descript.lower() == 'sunday' or work_descript.lower() == 'saturday':
            work_descript = 'regular day'
        else:
            work_descript = 'Special Holiday'

        # Work Hours Inputs

        time_in1 = datetime.strptime(str(row['Time In']), "%H:%M:%S").time()
        time_in = time_in1.strftime("%H:%M")
        time_out1 = datetime.strptime(str(row['Time Out']), "%H:%M:%S").time()
        time_out = time_out1.strftime("%H:%M")

        # Actual Time in and out
        actual_time_in1 = datetime.strptime(str(row['Actual Time In']), "%H:%M:%S").time()
        actual_time_out1 = datetime.strptime(str(row['Actual Time Out']), "%H:%M:%S").time()
        actual_time_in = actual_time_in1.strftime("%H:%M")
        actual_time_out = actual_time_out1.strftime("%H:%M")

        calculate_timeanddate(employee_name, employee_code, day_of_week, date_transact1, date_obj, work_descript,
                                  time_in, time_out, time_in1, time_out1, actual_time_in, actual_time_out,
                                  actual_time_in1, actual_time_out1)


    # Read the DataFrame
    df = pd.read_csv('dtr.csv')

    delete_duplicate(df)
    return render_template('index.html', dtr=dtr)

@app.route('/submit', methods=['POST'])
def submit():

    #Creating a new Dataframe (this serves as a database na rin)
    dtr = pd.DataFrame(columns=['Status',
                                'Overtime',
                                'Employee Code',
                                'Employee Name',
                                'Date',
                                'Day',
                                'Weekday or Weekend',
                                'Work Description',
                                'Time In',
                                'Time Out',
                                'Actual Time In',
                                'Actual Time Out',
                                'Net Hours Rendered (Time Format)',
                                'Actual Gross Hours Render',
                                'Hours Rendered',
                                'Undertime Hours',
                                'Tardiness',
                                'Excess of 8 hours Overtime',
                                'Total of 8 hours Overtime',
                                'Night Difference',
                                'Night Difference Overtime',
                                'Night Difference First 8 hours',
                                'Night Difference Excess of 8 hours'])
    #Employee inputs
    employee_name = request.form.get('employee_name')
    employee_code = request.form.get('employee_code')


    #Date of transaction
    date_transact = request.form.get('start_date')
    day_of_week = ""
    date_transact1 = ""
    if date_transact is not None:
        date_obj = datetime.strptime(date_transact, "%Y-%m-%d")
        date_transact1 = date_obj.date()
        day_of_week = date_obj.strftime("%A")
    else:
        date_obj = ""
        # Handle the case where the date is missing or empty


    # get the date value from the form submission
    date_str = request.form['start_date']

    # convert the date string to a datetime object
    date = datetime.strptime(date_str, '%Y-%m-%d')

    # add one day to the date value
    date += timedelta(days=1)

    # convert the modified date back to a string and store it in a variable
    modified_date_str = date.strftime('%Y-%m-%d')

    #Description inputs
    work_descript = request.form.get('work_descript').lower()

    #Work Hours Inputs
    time_in1 = datetime.strptime(request.form.get('time_in'), "%H:%M").time()
    time_in = time_in1.strftime("%H:%M")
    time_out1 = datetime.strptime(request.form.get('time_out'), "%H:%M").time()
    time_out = time_out1.strftime("%H:%M")

    #Actual Time in and out
    actual_time_in1 = datetime.strptime(request.form.get('actual_time_in'), "%H:%M").time()
    actual_time_in = actual_time_in1.strftime("%H:%M")
    actual_time_out1 = datetime.strptime(request.form.get('actual_time_out'), "%H:%M").time()
    actual_time_out = actual_time_out1.strftime("%H:%M")

    calculate_timeanddate(employee_name, employee_code, day_of_week, date_transact1, date_obj, work_descript, time_in, time_out, time_in1, time_out1, actual_time_in, actual_time_out, actual_time_in1, actual_time_out1)

    # Read the DataFrame
    df = pd.read_csv('dtr.csv')
    delete_duplicate(df)

    # Pass the HTML table to the template
    return render_template('index.html', dtr=dtr, modified_date_str=modified_date_str)

@app.route('/delete', methods=['POST'])
def delete_row():

    employee_id = request.form['delete_employee_code']
    date = request.form['delete_date']

    # Load the CSV file into a pandas dataframe
    df = pd.read_csv("dtr.csv")

    # Delete the row(s) based on the conditions
    if date == '':
        df = df.drop(df[(df['Employee Code'] == str(employee_id))].index)

    else:
        df = df.drop(df[(df['Employee Code'] == str(employee_id)) & (df['Date'] == date)].index)

    # Save the updated dataframe back to the CSV file
    df.to_csv("dtr.csv", index=False)
    delete_duplicate(df)
    data = df.sort_values(by=['Employee Name'])

    # Render the DataFrame as an HTML table
    html_table = data.to_html(index=False)

    return render_template('index.html', html_table=html_table)

@app.route('/table', methods=['GET', 'POST'])
def table():
    try:
        df = pd.read_csv('dtr.csv')
        data = df.to_dict('records')
        # add new table with summary information
        summary = pd.DataFrame({
            'Actual Gross Hours Rendered': [df['Actual Gross Hours Render'].sum()],
            'Hours Rendered': [df['Hours Rendered'].sum()],
            'Undertime Hours': [df['Undertime Hours'].sum()],
            'Tardiness': [df['Tardiness'].sum()],
            'Night Difference': [df['Night Difference'].sum()],
            'Night Difference Overtime': [df['Night Difference Overtime'].sum()]
        })
        summary_html = summary.to_html(index=False)
        return render_template('table.html', data=data, summary=summary_html)
    except FileNotFoundError:
        return render_template('table.html')


@app.route('/save_row', methods=['POST'])
def save_row():
    index = int(request.form['index'])
    status = request.form['status']
    df = pd.read_csv('dtr.csv')
    df.loc[index, 'Status'] = status
    df.to_csv('dtr.csv', index=False)
    return redirect(url_for('table'))

@app.route('/download')
def download():
    # Step 1: Read input file into a DataFrame
    global sum_days, new_df, employee_code, name, RegularDay, RegularDay_RestDay, RegularDay_Overtime, RegularDay_RestDay_Overtime, RegularDay_Night_diffence, RegularDay_Night_diffence_Overtime, RegularDay_Night_diffence_Restday, RegularDay_Night_diffence_Restday_Overtime, Legal_Holiday, Legal_Holiday_Overtime, Legal_Holiday_RestDay, Legal_Holiday_RestDay_Overtime, Legal_Holiday_Night_diffence, Legal_Holiday_Night_diffence_Overtime, Legal_Holiday_Night_diffence_Restday, Legal_Holiday_Night_diffence_Restday_Overtime, Special_Holiday, Special_Holiday_Overtime, Special_Holiday_RestDay, Special_Holiday_RestDay_Overtime, Special_Holiday_Night_diffence, Special_Holiday_Night_diffence_Overtime, Special_Holiday_Night_diffence_Restday, Special_Holiday_Night_diffence_Restday_Overtime

    try:
        df = pd.read_csv('dtr.csv')
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Employee_Code',
                                   'Employee_Name',
                                   'No. of Working Days',
                                   'Number of Working Hours',
                                   'Tardiness/Undertime',
                                   'No. of Days Absent',
                                   'ROT_125',
                                   'Regular OT 100',
                                   'Rest Day',
                                   'RestDay Overtime for the 1st 8hrs',
                                   'Rest Day Overtime in Excess of 8hrs',
                                   'Special Holiday',
                                   'Special Holiday_1st 8hrs',
                                   'Special Holiday_Excess of 8hrs',
                                   'Special Holiday Falling on restday 1st 8hrs',
                                   'Legal Holiday',
                                   'Legal Holiday_1st 8hrs',
                                   'Legal Holiday_Excess of 8hrs',
                                   'Legal Holiday Falling on Rest Day_1st 8hrs',
                                   'Legal Holiday Falling on Rest Day_Excess of 8hrs',
                                   'Night Differential Regular Days_1st 8hrs',
                                   'Night Differential Regular Days_Excess of 8hrs',
                                   'Night Differential Falling on Rest Day_1st 8hrs',
                                   'Night Differential Falling on Rest Day_Excess of 8hrs',
                                   'Night Differential Falling on SPHOL rest day 1st 8 hr',
                                   'Night Differential SH falling on RD_EX8',
                                   'Night Differential on Legal Holidays falling on Rest Days',
                                   'Night Differential on Legal Holidays_1st 8hrs',
                                   'Night Differential on Legal Holidays_Excess of 8hrs',
                                   'Night Differential falling on Special Holiday',
                                   'Night Differential SH_EX8'])


    file_name = 'DTR_Summary.csv'
    if os.path.exists(file_name):
        os.remove(file_name)


    # Step 4: Compute overtime for each employee
    for code in df['Employee Name'].unique():
        employee_df = df[df['Employee Name'] == code]
        name = employee_df['Employee Name'].iloc[0]
        employee_code = employee_df['Employee Code'].iloc[0]
        status = employee_df['Status']



        # Check if the status is not approved and set overtime to 0
        # if 'Not Approved' in status.values:
        #     employee_df.loc[status == 'Not Approved', 'Overtime'] = 0

        total_working_days = 0
        working_hours1 = 0
        tardiness1 = 0
        RegularDay_RestDay = 0
        RegularDay_Overtime = 0
        RegularDay_Overtime_Excess = 0
        RegularDay_RestDay_Overtime = 0
        RegularDay_RestDay_Overtime_Excess = 0

        RegularDay_RestDay_Night_8hours = 0
        RegularDay_RestDay_Night_8hours_Excess = 0
        RegularDay_Night_8hours = 0
        RegularDay_Night_8hours_Excess = 0

        Specialholiday = 0
        Specialholiday_Overtime = 0
        Specialholiday_Overtime_Excess = 0
        Specialholiday_RestDay_Overtime = 0
        legalholiday = 0
        legalholiday_Overtime = 0
        legalholiday_Overtime_Excess = 0
        legalholiday_RestDay_Overtime = 0
        legalholiday_RestDay_Overtime_Excess = 0
        Specialholiday_RestDay_Night_8hours = 0
        Specialholiday_RestDay_Night_8hours_Excess = 0
        legalholiday_Night_diffence = 0
        legalholiday_Night_8hours = 0
        legalholiday_Night_8hours_Excess = 0
        Specialholiday_Night_diffence = 0
        Specialholiday_Night_8hours_Excess = 0

        for index, row in employee_df.iterrows():
            work_desc = row['Work Description']
            overtime = row['Overtime']
            rest_day = row['Weekday or Weekend']
            time_in = row['Actual Time In']
            working_hours = row['Actual Gross Hours Render']
            tardiness = row['Tardiness']

            if time_in is not None:
                total_working_days = total_working_days + 1

            working_hours1 = working_hours + working_hours1

            tardiness1 = tardiness + tardiness1

            if work_desc == 'regular day':
                if rest_day == 'Weekend':
                    RegularDay_RestDay = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Actual Gross Hours Render'].sum()
                    RegularDay_Night_diffence_Restday = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference'].sum()


                    if overtime >= 1:
                        RegularDay_RestDay_Overtime1 = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Overtime'].sum()
                        RegularDay_Night_diffence_Restday_Overtime = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference Overtime'].sum()
                        RegularDay_Night_diffence_Restday = RegularDay_Night_diffence_Restday - RegularDay_Night_diffence_Restday_Overtime
                        RegularDay_RestDay_Overtime_Excess = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Excess of 8 hours Overtime'].sum()
                        RegularDay_RestDay_Overtime = RegularDay_RestDay_Overtime1 - RegularDay_RestDay_Overtime_Excess


                        if RegularDay_RestDay_Overtime1 > RegularDay_RestDay:
                            RegularDay_RestDay = RegularDay_RestDay_Overtime1 - RegularDay_RestDay

                        elif RegularDay_RestDay_Overtime1 < RegularDay_RestDay:
                            RegularDay_RestDay = RegularDay_RestDay - RegularDay_RestDay_Overtime1

                        RegularDay_RestDay_Night_8hours = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference First 8 hours'].sum()
                        RegularDay_RestDay_Night_8hours_Excess = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference Excess of 8 hours'].sum()


                elif rest_day == 'Weekday':
                    RegularDay = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Actual Gross Hours Render'].sum()
                    RegularDay_Night_diffence = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference'].sum()

                    if overtime >= 1:
                        RegularDay_Overtime1 = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Overtime'].sum()
                        RegularDay_Night_diffence_Overtime = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference Overtime'].sum()
                        RegularDay_Night_diffence = RegularDay_Night_diffence - RegularDay_Night_diffence_Overtime
                        RegularDay_Overtime_Excess = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Excess of 8 hours Overtime'].sum()
                        RegularDay_Overtime = RegularDay_Overtime1 - RegularDay_Overtime_Excess
                        
                        if RegularDay_Overtime1 > RegularDay:
                            RegularDay = RegularDay_Overtime1 - RegularDay

                        elif RegularDay_Overtime1 < RegularDay:
                            RegularDay = RegularDay - RegularDay_Overtime1

                        RegularDay_Night_8hours = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference First 8 hours'].sum()
                        RegularDay_Night_8hours_Excess = employee_df.loc[(employee_df['Work Description'] == 'regular day') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference Excess of 8 hours'].sum()

            if work_desc == 'special holiday':
                if rest_day == 'Weekend':
                    Specialholiday_RestDay = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Actual Gross Hours Render'].sum()
                    Specialholiday_Night_diffence_Restday = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference'].sum()

                    if overtime >= 1:
                        Specialholiday_RestDay_Overtime1 = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Overtime'].sum()
                        Specialholiday_Night_diffence_Restday_Overtime = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference Overtime'].sum()
                        Specialholiday_Night_diffence_Restday = Specialholiday_Night_diffence_Restday - Specialholiday_Night_diffence_Restday_Overtime
                        Specialholiday_RestDay_Overtime_Excess = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Excess of 8 hours Overtime'].sum()
                        Specialholiday_RestDay_Overtime = Specialholiday_RestDay_Overtime1 - Specialholiday_RestDay_Overtime_Excess

                        if Specialholiday_RestDay_Overtime1 > Specialholiday_RestDay:
                            Specialholiday_RestDay = Specialholiday_RestDay_Overtime1 - Specialholiday_RestDay

                        elif Specialholiday_RestDay_Overtime1 < Specialholiday_RestDay:
                            Specialholiday_RestDay = Specialholiday_RestDay - Specialholiday_RestDay_Overtime1

                        Specialholiday_RestDay_Night_8hours = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference First 8 hours'].sum()
                        Specialholiday_RestDay_Night_8hours_Excess = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference Excess of 8 hours'].sum()


                elif rest_day == 'Weekday':
                    Specialholiday = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Actual Gross Hours Render'].sum()
                    Specialholiday_Night_diffence = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference'].sum()

                    if overtime >= 1:
                        Specialholiday_Overtime1 = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Overtime'].sum()
                        Specialholiday_Night_diffence_Overtime = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference Overtime'].sum()
                        Specialholiday_Night_diffence = Specialholiday_Night_diffence - Specialholiday_Night_diffence_Overtime
                        Specialholiday_Overtime_Excess = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Excess of 8 hours Overtime'].sum()
                        Specialholiday_Overtime = Specialholiday_Overtime1 - Specialholiday_Overtime_Excess

                        if Specialholiday_Overtime1 > Specialholiday:
                            Specialholiday = Specialholiday_Overtime1 - Specialholiday

                        elif Specialholiday_Overtime1 < Specialholiday:
                            Specialholiday = Specialholiday - Specialholiday_Overtime1

                        Specialholiday_Night_8hours = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference First 8 hours'].sum()
                        Specialholiday_Night_8hours_Excess = employee_df.loc[(employee_df['Work Description'] == 'special holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference Excess of 8 hours'].sum()

            if work_desc == 'legal holiday':
                if rest_day == 'Weekend':
                    legalholiday_RestDay = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Actual Gross Hours Render'].sum()
                    legalholiday_Night_diffence_Restday = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference'].sum()

                    if overtime >= 1:
                        legalholiday_RestDay_Overtime1 = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Overtime'].sum()
                        legalholiday_Night_diffence_Restday_Overtime = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference Overtime'].sum()
                        legalholiday_Night_diffence_Restday = legalholiday_Night_diffence_Restday - legalholiday_Night_diffence_Restday_Overtime
                        legalholiday_RestDay_Overtime_Excess = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Excess of 8 hours Overtime'].sum()
                        legalholiday_RestDay_Overtime = legalholiday_RestDay_Overtime1 - legalholiday_RestDay_Overtime_Excess

                        if legalholiday_RestDay_Overtime1 > legalholiday_RestDay:
                            legalholiday_RestDay = legalholiday_RestDay_Overtime1 - legalholiday_RestDay

                        elif legalholiday_RestDay_Overtime1 < legalholiday_RestDay:
                            legalholiday_RestDay = legalholiday_RestDay - legalholiday_RestDay_Overtime1

                        legalholiday_RestDay_Night_8hours = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference First 8 hours'].sum()
                        legalholiday_RestDay_Night_8hours_Excess = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekend'), 'Night Difference Excess of 8 hours'].sum()


                elif rest_day == 'Weekday':
                    legalholiday = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Actual Gross Hours Render'].sum()
                    legalholiday_Night_diffence = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference'].sum()

                    if overtime >= 1:
                        legalholiday_Overtime1 = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Overtime'].sum()
                        legalholiday_Night_diffence_Overtime = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df[
                                                                                      'Weekday or Weekend'] == 'Weekday'), 'Night Difference Overtime'].sum()
                        legalholiday_Night_diffence = legalholiday_Night_diffence - legalholiday_Night_diffence_Overtime
                        legalholiday_Overtime_Excess = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Excess of 8 hours Overtime'].sum()
                        legalholiday_Overtime = legalholiday_Overtime1 - legalholiday_Overtime_Excess

                        if legalholiday_Overtime1 > legalholiday:
                            legalholiday = legalholiday_Overtime1 - legalholiday

                        elif legalholiday_Overtime1 < legalholiday:
                            legalholiday = legalholiday - legalholiday_Overtime1

                        legalholiday_Night_8hours = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference First 8 hours'].sum()
                        legalholiday_Night_8hours_Excess = employee_df.loc[(employee_df['Work Description'] == 'legal holiday') & (employee_df['Weekday or Weekend'] == 'Weekday'), 'Night Difference Excess of 8 hours'].sum()

        new_df = pd.DataFrame({'Employee_Code': [employee_code],
                               'Employee_Name': [name],
                               'No. of Working Days': [total_working_days],
                               'Number of Working Hours': [working_hours1],
                               'Tardiness/Undertime': [tardiness1],
                               'ROT_125': [RegularDay_Overtime],
                               'Regular OT 100': [RegularDay_Overtime_Excess],
                               'Rest Day': [RegularDay_RestDay],
                               'RestDay Overtime for the 1st 8hrs': [RegularDay_RestDay_Overtime],
                               'Rest Day Overtime in Excess of 8hrs': [RegularDay_RestDay_Overtime_Excess],
                               'Special Holiday':[Specialholiday],
                                'Special Holiday_1st 8hrs':[Specialholiday_Overtime],
                                'Special Holiday_Excess of 8hrs':[Specialholiday_Overtime_Excess],
                                'Special Holiday Falling on restday 1st 8hrs':[Specialholiday_RestDay_Overtime],
                                'Legal Holiday':[legalholiday],
                                'Legal Holiday_1st 8hrs':[legalholiday_Overtime],
                                'Legal Holiday_Excess of 8hrs':[legalholiday_Overtime_Excess],
                                'Legal Holiday Falling on Rest Day_1st 8hrs':[legalholiday_RestDay_Overtime],
                                'Legal Holiday Falling on Rest Day_Excess of 8hrs':[legalholiday_RestDay_Overtime_Excess],
                                'Night Differential Regular Days_1st 8hrs':[RegularDay_Night_8hours],
                                'Night Differential Regular Days_Excess of 8hrs':[RegularDay_Night_8hours_Excess],
                                'Night Differential Falling on Rest Day_1st 8hrs':[RegularDay_RestDay_Night_8hours],
                                'Night Differential Falling on Rest Day_Excess of 8hrs':[RegularDay_RestDay_Night_8hours_Excess],
                                'Night Differential Falling on SPHOL rest day 1st 8 hr':[Specialholiday_RestDay_Night_8hours],
                                'Night Differential SH falling on RD_EX8':[Specialholiday_RestDay_Night_8hours_Excess],
                                'Night Differential on Legal Holidays falling on Rest Days':[legalholiday_Night_diffence],
                                'Night Differential on Legal Holidays_1st 8hrs':[legalholiday_Night_8hours],
                                'Night Differential on Legal Holidays_Excess of 8hrs':[legalholiday_Night_8hours_Excess],
                                'Night Differential falling on Special Holiday':[Specialholiday_Night_diffence],
                                'Night Differential SH_EX8':[Specialholiday_Night_8hours_Excess]})

        try:
            existing_df = pd.read_csv('DTR_Summary.csv')
        except FileNotFoundError:
            existing_df = pd.DataFrame(columns=['Employee_Code',
                                                'Employee_Name',
                                                'No. of Working Days',
                                                'Number of Working Hours',
                                                'Tardiness/Undertime',
                                                'No. of Days Absent',
                                                'ROT_125',
                                                'Regular OT 100',
                                                'Rest Day',
                                                'RestDay Overtime for the 1st 8hrs',
                                                'Rest Day Overtime in Excess of 8hrs',
                                                'Special Holiday',
                                                'Special Holiday_1st 8hrs',
                                                'Special Holiday_Excess of 8hrs',
                                                'Special Holiday Falling on restday 1st 8hrs',
                                                'Legal Holiday',
                                                'Legal Holiday_1st 8hrs',
                                                'Legal Holiday_Excess of 8hrs',
                                                'Legal Holiday Falling on Rest Day_1st 8hrs',
                                                'Legal Holiday Falling on Rest Day_Excess of 8hrs',
                                                'Night Differential Regular Days_1st 8hrs',
                                                'Night Differential Regular Days_Excess of 8hrs',
                                                'Night Differential Falling on Rest Day_1st 8hrs',
                                                'Night Differential Falling on Rest Day_Excess of 8hrs',
                                                'Night Differential Falling on SPHOL rest day 1st 8 hr',
                                                'Night Differential SH falling on RD_EX8',
                                                'Night Differential on Legal Holidays falling on Rest Days',
                                                'Night Differential on Legal Holidays_1st 8hrs',
                                                'Night Differential on Legal Holidays_Excess of 8hrs',
                                                'Night Differential falling on Special Holiday',
                                                'Night Differential SH_EX8'
                                                ])

        new_df = pd.concat([existing_df, new_df], ignore_index=True)
        add_df = new_df.sort_values(by=['Employee_Name'])
        add_df.to_csv('DTR_Summary.csv', index=False)

    try:
        csv_excel = pd.read_csv('DTR_Summary.csv')
    except FileNotFoundError:
        csv_excel = pd.DataFrame(columns=['Employee_Code',
                                            'Employee_Name',
                                            'Regular_Day',
                                            'RegularDay_Overtime',
                                            'RegularDay_RestDay',
                                            'RegularDay_RestDay_Overtime',
                                            'RegularDay_Night_diffence',
                                            'RegularDay_Night_diffence_Overtime',
                                            'RegularDay_Night_diffence_Restday',
                                            'RegularDay_Night_diffence_Restday_Overtime',
                                            'Legal_Holiday',
                                            'Legal_Holiday_Overtime',
                                            'Legal_Holiday_RestDay',
                                            'Legal_Holiday_RestDay_Overtime',
                                            'Legal_Holiday_Night_diffence',
                                            'Legal_Holiday_Night_diffence_Overtime',
                                            'Legal_Holiday_Night_diffence_Restday',
                                            'Legal_Holiday_Night_diffence_Restday_Overtime',
                                            'Special_Holiday',
                                            'Special_Holiday_Overtime',
                                            'Special_Holiday_RestDay',
                                            'Special_Holiday_RestDay_Overtime',
                                            'Special_Holiday_Night_diffence',
                                            'Special_Holiday_Night_diffence_Overtime',
                                            'Special_Holiday_Night_diffence_Restday',
                                            'Special_Holiday_Night_diffence_Restday_Overtime'])
    csv_excel.to_excel('DTR_Summary.xlsx', index=False)

    # Download the new CSV file
    return send_file('DTR_Summary.xlsx', as_attachment=True)



if __name__ == '__main__':
    app.run(use_reloader=True, port=port)
