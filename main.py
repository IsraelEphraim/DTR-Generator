import os
from flask import Flask, render_template, request, redirect, send_file, jsonify, url_for
import pandas as pd
from datetime import datetime, timedelta, time
import webview
import sys

base_dir = '.'
if hasattr(sys, '_MEIPASS'):
    base_dir = os.path.join(sys._MEIPASS)

port = int(os.environ.get('PORT', 9000))

app = Flask(__name__, template_folder=os.path.join(base_dir, 'template'))
window = webview.create_window('DTR Generator', app)

def delete_duplicate(df):
    df.drop_duplicates(subset=['Employee Code', 'Date'], inplace=True)
    df.to_csv('dtr.csv', index=False)
    return df

@app.route('/')
def index():
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
                                    'RestDay Overtime for the 1st 8hrs',
                                    'Rest Day Overtime in Excess of 8hrs',
                                    'Special Holiday',
                                    'Special Holiday_1st 8hours',
                                    'Special Holiday_Excess of 8hrs',
                                    'Special Holiday Falling on restday 1st 8hrs',
                                    'Special Holiday on restday Excess 8Hrs',
                                    'Legal Holiday',
                                    'Legal Holiday_1st 8hours',
                                    'Legal Holiday_Excess of 8hrs',
                                    'Legal Holiday Falling on Rest Day_1st 8hrs',
                                    'Legal Holiday Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential Regular Days_1st 8hrs',
                                    'Night Differential Regular Days_Excess of 8hrs',
                                    'Night Differential Falling on Rest Day_1st 8hrs',
                                    'Night Differential Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential falling on Special Holiday',
                                    'Night Differential SH_EX8',
                                    'Night Differential Falling on SPHOL rest day 1st 8 hr',
                                    'Night Differential SH falling on RD_EX8',
                                    'Night Differential on Legal Holidays_1st 8hrs',
                                    'Night Differential on Legal Holidays_Excess of 8hrs',
                                    'Night Differential on Legal Holidays falling on Rest Days'
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
                                    'RestDay Overtime for the 1st 8hrs',
                                    'Rest Day Overtime in Excess of 8hrs',
                                    'Special Holiday',
                                    'Special Holiday_1st 8hours',
                                    'Special Holiday_Excess of 8hrs',
                                    'Special Holiday Falling on restday 1st 8hrs',
                                    'Special Holiday on restday Excess 8Hrs',
                                    'Legal Holiday',
                                    'Legal Holiday_1st 8hours',
                                    'Legal Holiday_Excess of 8hrs',
                                    'Legal Holiday Falling on Rest Day_1st 8hrs',
                                    'Legal Holiday Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential Regular Days_1st 8hrs',
                                    'Night Differential Regular Days_Excess of 8hrs',
                                    'Night Differential Falling on Rest Day_1st 8hrs',
                                    'Night Differential Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential falling on Special Holiday',
                                    'Night Differential SH_EX8',
                                    'Night Differential Falling on SPHOL rest day 1st 8 hr',
                                    'Night Differential SH falling on RD_EX8',
                                    'Night Differential on Legal Holidays_1st 8hrs',
                                    'Night Differential on Legal Holidays_Excess of 8hrs',
                                    'Night Differential on Legal Holidays falling on Rest Days'])

    return render_template('index.html', dtr=dtr)

#Auto Generate of Name using Employee Code
@app.route("/lookup_employee", methods=["POST"])
def lookup_employee():
    employee_data = pd.read_csv("dtr.csv")
    employee_code = request.json["employee_code"]
    employee_name = employee_data.loc[employee_data["Employee Code"] == employee_code, "Employee Name"].iloc[0]
    return jsonify({"employee_name": employee_name})

def timedelta_to_decimal(timedelta_value):
    total_seconds = timedelta_value.total_seconds()
    decimal_value = total_seconds / 3600  # Divide by 3600 to convert seconds to decimal hours
    return decimal_value

def hour_estimate(estimate):

    x = int(estimate)
    y = estimate - int(estimate)
    if y > .50:
        y = .50
        estimate = x + y

    elif y < .50:
        y = .0
        estimate = x + y

    if estimate < 1:
        estimate = 0

    return estimate

def calculate_timeanddate(employee_name, employee_code, day_of_week, date_transact1, date_obj, work_descript, time_in, time_out, time_in1, time_out1, actual_time_in, actual_time_out, actual_time_in1, actual_time_out1):
    week_day = date_obj.weekday()
    if week_day < 5:  # Weekday (Monday to Friday)
        week_check = "Weekday"
    else:  # Weekend day (Saturday or Sunday)
        week_check = "Weekend"


    # =====================================================================
    #TOTAL OF SCHEDULED TIME IN AND TIME OUT
    datetime_timein = datetime.combine(date_obj.date(), time_in1)
    if time_in1 > time_out1:
        add_day = timedelta(days=1)
        new_day = date_obj.date() + add_day
        datetime_timeout = datetime.combine(new_day, time_out1)

    else:
        datetime_timeout = datetime.combine(date_obj.date(), time_out1)

    #ITO YUNG TOTAL BOSS
    total_datetime_in_out = datetime_timeout - datetime_timein
    total_datetime_in_out_int = timedelta_to_decimal(total_datetime_in_out)

    #=====================================================================
    # TOTAL OF ACTUAL TIME IN AND TIME OUT
    actual_datetime_timein = datetime.combine(date_obj.date(), actual_time_in1)
    if actual_time_in1 > actual_time_out1:
        actual_add_day = timedelta(days=1)
        actual_new_day = date_obj.date() + actual_add_day
        actual_datetime_timeout = datetime.combine(actual_new_day, actual_time_out1)

    else:
        actual_datetime_timeout = datetime.combine(date_obj.date(), actual_time_out1)

    # ITO YUNG TOTAL BOSS
    total_actual_datetime_in_out =  actual_datetime_timeout - actual_datetime_timein
    total_actual_datetime_in_out_int = timedelta_to_decimal(total_actual_datetime_in_out)
    print(total_actual_datetime_in_out_int)
    #=====================================================================
    #NON CHARGEABLE BREAK
    break_start = '12:00'
    break_end = '13:00'
    break_start_convert = datetime.strptime(break_start, "%H:%M").time()
    break_end_convert = datetime.strptime(break_end, "%H:%M").time()
    break_start_convert_1 = datetime.combine(date_obj, break_start_convert)
    break_end_convert_1 = datetime.combine(date_obj, break_end_convert)

    if actual_datetime_timein < break_start_convert_1 and actual_datetime_timeout > break_end_convert_1:
        non_chargeable_break = time(1, 0, 0)
        non_chargeable_break_timedelta = timedelta(hours=non_chargeable_break.hour, minutes=non_chargeable_break.minute,
                                                   seconds=non_chargeable_break.second)

    else:
        non_chargeable_break = time(0, 0, 0)
        non_chargeable_break_timedelta = timedelta(hours=non_chargeable_break.hour, minutes=non_chargeable_break.minute,
                                                   seconds=non_chargeable_break.second)

    #=====================================================================
    #NET HOURS RENDERED
    if week_check == 'Weekday' and work_descript == 'regular day':
        net_hours_rendered = total_datetime_in_out - non_chargeable_break_timedelta
        net_hours_rendered_str = timedelta_to_decimal(net_hours_rendered)

    else:
        net_hours_rendered_str = 0

    #=====================================================================
    #ACTUAL GROSS RENDERED

    actual_render = 0
    if total_actual_datetime_in_out_int > 0:
        actual_render = total_actual_datetime_in_out - non_chargeable_break_timedelta
        actual_render = timedelta_to_decimal(actual_render)
        actual_render = round(actual_render, 2)

    elif work_descript == 'legal holiday' and total_actual_datetime_in_out_int == 0:
        actual_render = 0

    elif work_descript == 'special holiday' and total_actual_datetime_in_out_int == 0:
        actual_render = 0

    #=====================================================================
    #HOURS RENDERED

    hour_rendered = total_actual_datetime_in_out
    hour_rendered_str = timedelta_to_decimal(hour_rendered)
    hour_rendered_str = int(hour_rendered_str)

    if work_descript == 'legal holiday' or work_descript == 'special holiday' and total_actual_datetime_in_out_int == 0 and total_datetime_in_out_int == 0:
        hour_rendered_str = 0


    if hour_rendered_str > 8:
        hour_rendered_str = 8


    #=====================================================================
    #OVERTIME

    #NGHT DIFFERENCE LODS
    night_diff_start = '22:00'
    night_diff_end = '06:00'

    datetime_start_night = datetime.strptime(night_diff_start, "%H:%M").time()
    datetime_start_night = datetime.combine(date_obj, datetime_start_night)
    datetime_end_night = datetime.strptime(night_diff_end, "%H:%M").time()
    one_day = timedelta(days=1)
    new_date = date_obj + one_day
    date_time_end_night = datetime.combine(new_date, datetime_end_night)


    #VARIABLES TO LODS
    overtime = 0
    ot_excess_8 = 0
    ot_first_8  = 0
    restday_ot_first_8 = 0
    restday_ot_excess_8 = 0
    legal_holiday = 0
    special_holiday = 0
    legal_holiday_excess_8 = 0
    legal_holiday_1st_8 = 0
    legal_holiday_rd_excess_8 = 0
    legal_holiday_rd_1st_8 = 0
    special_holiday_excess_8 = 0
    special_holiday_1st_8 = 0
    special_holiday_rd_excess_8 = 0
    special_holiday_rd_1st_8 = 0
    ND_regular_days_1st_8_hrs = 0
    ND_regular_days_excess_8_hrs = 0
    ND_regular_days_RD_1st_8_hrs = 0
    ND_regular_days_RD_excess_8_hrs = 0

    if week_check == 'Weekday' and work_descript == 'regular day':
        if actual_render > 8:
            overtime = actual_render - 8
            if overtime > 8:
                ot_excess_8 = overtime - 8
                ot_first_8 = overtime - ot_excess_8
                ot_excess_8 = round(ot_excess_8, 2)
                ot_first_8 = round(ot_first_8, 2)

            elif overtime < 8:
                ot_first_8 = overtime
                ot_first_8 = round(ot_first_8, 2)

            # NIGHT DIFFERENCE OT EXCESS 8
            if actual_datetime_timeout > datetime_start_night and actual_datetime_timeout < date_time_end_night:

                actual_night_out = actual_datetime_timeout - datetime_start_night
                night_out = datetime_start_night - datetime_timeout
                night_out = timedelta_to_decimal(night_out)

                if night_out < 8:
                    ND_regular_days_1st_8_hrs = ot_first_8 - night_out

                elif night_out > 8:
                    ND_regular_days_1st_8_hrs = ot_first_8 - night_out
                    ND_regular_days_excess_8_hrs = ot_excess_8 - ND_regular_days_1st_8_hrs

            ot_first_8 = hour_estimate(ot_first_8)
            ot_excess_8 = hour_estimate(ot_excess_8)

        overtime = hour_estimate(overtime)



    elif week_check == 'Weekend' and work_descript == 'regular day':
        if actual_render > 8:
            restday_ot_excess_8 = actual_render - 8
            restday_ot_first_8 = actual_render - restday_ot_excess_8
            restday_ot_excess_8 = round(restday_ot_excess_8, 2)
            overtime = 0

            # NIGHT DIFFERENCE OT EXCESS 8
            if actual_datetime_timeout > datetime_start_night and actual_datetime_timeout < date_time_end_night:

                actual_night_out = actual_datetime_timeout - datetime_start_night
                night_out = datetime_start_night - datetime_timeout
                night_out = timedelta_to_decimal(night_out)

                if night_out < 8:
                    ND_regular_days_RD_1st_8_hrs = restday_ot_first_8 - night_out
                    ND_regular_days_RD_excess_8_hrs = restday_ot_excess_8 - ND_regular_days_RD_1st_8_hrs


            restday_ot_first_8 = hour_estimate(restday_ot_first_8)
            restday_ot_excess_8 = hour_estimate(restday_ot_excess_8)

        elif actual_render <= 8:
            restday_ot_first_8 = actual_render
            overtime = 0
            restday_ot_first_8 = hour_estimate(restday_ot_first_8)

            #NIGHT DIFFERENCE
            if actual_datetime_timeout > datetime_start_night and actual_datetime_timeout < date_time_end_night:
                ND_regular_days_RD_1st_8_hrs = actual_datetime_timeout - datetime_start_night
                ND_regular_days_RD_1st_8_hrs1 = restday_ot_first_8 - ND_regular_days_RD_1st_8_hrs
                ND_regular_days_RD_1st_8_hrs = ND_regular_days_RD_1st_8_hrs - ND_regular_days_RD_1st_8_hrs1
                ND_regular_days_RD_1st_8_hrs = timedelta_to_decimal(ND_regular_days_RD_1st_8_hrs)
                ND_regular_days_RD_1st_8_hrs = hour_estimate(ND_regular_days_RD_1st_8_hrs)




    #IF DI PUMASOK YUNG EMPLOYEE
    elif work_descript == 'legal holiday' and total_actual_datetime_in_out_int == 0 and total_datetime_in_out_int == 0:
        overtime = 0
        legal_holiday = 8

    elif work_descript == 'special holiday' and total_actual_datetime_in_out_int == 0 and total_datetime_in_out_int == 0:
        overtime = 0
        special_holiday = 8

    #IF PUMASOK YUNG EMPLOYEE DITO MAGCOCOMPUTE
    elif work_descript == 'legal holiday' and total_actual_datetime_in_out_int > 0 and total_datetime_in_out_int > 0:
        if week_check == 'Weekday':
            if actual_render > 8:
                legal_holiday_excess_8 = actual_render - 8
                legal_holiday_1st_8 = actual_render - legal_holiday_excess_8
                legal_holiday_excess_8 = round(legal_holiday_excess_8, 2)



            elif actual_render <= 8:
                legal_holiday_1st_8 = actual_render
                legal_holiday_1st_8 = round(legal_holiday_1st_8, 2)



        elif week_check == 'Weekend':
            if actual_render > 8:
                legal_holiday_rd_excess_8 = actual_render - 8
                legal_holiday_rd_1st_8 = actual_render - legal_holiday_rd_excess_8
                legal_holiday_rd_excess_8 = round(legal_holiday_rd_excess_8, 2)


            elif actual_render <= 8:

                legal_holiday_rd_1st_8 = actual_render
                legal_holiday_rd_1st_8 = round(legal_holiday_rd_1st_8, 2)

        legal_holiday_1st_8 = hour_estimate(legal_holiday_1st_8)
        legal_holiday_excess_8 = hour_estimate(legal_holiday_excess_8)
        legal_holiday_rd_1st_8 = hour_estimate(legal_holiday_rd_1st_8)
        legal_holiday_rd_excess_8 = hour_estimate(legal_holiday_rd_excess_8)

    elif work_descript == 'special holiday' and total_actual_datetime_in_out_int > 0 and total_datetime_in_out_int > 0:
        if week_check == 'Weekday':
            if actual_render > 8:
                special_holiday_excess_8 = actual_render - 8
                special_holiday_1st_8 = actual_render - special_holiday_excess_8
                special_holiday_excess_8 = round(special_holiday_excess_8, 2)

            elif actual_render <= 8:
                special_holiday_1st_8 = actual_render
                special_holiday_1st_8 = round(special_holiday_1st_8, 2)
        elif week_check == 'Weekend':
            if actual_render > 8:
                special_holiday_rd_excess_8 = actual_render - 8
                special_holiday_rd_1st_8 = actual_render - special_holiday_rd_excess_8
                special_holiday_rd_excess_8 = round(special_holiday_rd_excess_8, 2)

            elif actual_render <= 8:
                special_holiday_rd_1st_8 = actual_render
                special_holiday_rd_1st_8 = round(special_holiday_rd_1st_8, 2)

            special_holiday_excess_8 = hour_estimate(special_holiday_excess_8)
            special_holiday_1st_8 = hour_estimate(special_holiday_1st_8)
            special_holiday_rd_excess_8 = hour_estimate(special_holiday_rd_excess_8)
            special_holiday_rd_1st_8 = hour_estimate(special_holiday_rd_1st_8)



    #=====================================================================
    #UNDERTIME HOURS

    if week_check == 'Weekday' and work_descript == 'regular day':
        if total_actual_datetime_in_out < total_datetime_in_out:
            undertime_check_try = total_datetime_in_out - total_actual_datetime_in_out
            undertime_check_try = timedelta_to_decimal(undertime_check_try)
        else:
            undertime_check_try = 0

    else:
        undertime_check_try = 0

    #=====================================================================
    #TARDINESS

    tardiness_str = 0
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

    #=====================================================================
    # Approval
    status_OT = 'No OT'
    if overtime >= 1:
        status_OT = 'Not Approved'
    elif work_descript == 'regular day' and overtime == 0:
        status_OT = 'No OT'
    elif work_descript == 'legal holiday' or work_descript == 'special holiday' and overtime == 0:
        status_OT = 'Holiday'

    # Encoding of inputs to the dataframe
    dtr_new = pd.DataFrame({'Status': [status_OT],
                            'Overtime': [overtime],
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
                            'Actual Gross Hours Render': [actual_render],
                            'Hours Rendered': [hour_rendered_str],
                            'Undertime Hours': [undertime_check_try],
                            'Tardiness': [tardiness_str],
                            'Excess of 8 hours Overtime': [ot_excess_8],
                            'Total of 8 hours Overtime': [ot_first_8],
                            'RestDay Overtime for the 1st 8hrs': [restday_ot_first_8],
                            'Rest Day Overtime in Excess of 8hrs': [restday_ot_excess_8],
                            'Special Holiday': [special_holiday],
                            'Special Holiday_1st 8hours': [special_holiday_1st_8],
                            'Special Holiday_Excess of 8hrs': [special_holiday_excess_8],
                            'Special Holiday Falling on restday 1st 8hrs': [special_holiday_rd_1st_8],
                            'Special Holiday on restday Excess 8Hrs': [special_holiday_rd_excess_8],
                            'Legal Holiday': [legal_holiday],
                            'Legal Holiday_1st 8hours': [legal_holiday_1st_8],
                            'Legal Holiday_Excess of 8hrs': [legal_holiday_excess_8],
                            'Legal Holiday Falling on Rest Day_1st 8hrs': [legal_holiday_rd_1st_8],
                            'Legal Holiday Falling on Rest Day_Excess of 8hrs': [legal_holiday_rd_excess_8],
                            'Night Differential Regular Days_1st 8hrs': [ND_regular_days_1st_8_hrs],
                            'Night Differential Regular Days_Excess of 8hrs': [ND_regular_days_excess_8_hrs],
                            'Night Differential Falling on Rest Day_1st 8hrs': [ND_regular_days_RD_1st_8_hrs],
                            'Night Differential Falling on Rest Day_Excess of 8hrs': [ND_regular_days_RD_excess_8_hrs]
                            })

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
                                    'RestDay Overtime for the 1st 8hrs',
                                    'Rest Day Overtime in Excess of 8hrs',
                                    'Special Holiday',
                                    'Special Holiday_1st 8hours',
                                    'Special Holiday_Excess of 8hrs',
                                    'Special Holiday Falling on restday 1st 8hrs',
                                    'Special Holiday on restday Excess 8Hrs',
                                    'Legal Holiday',
                                    'Legal Holiday_1st 8hours',
                                    'Legal Holiday_Excess of 8hrs',
                                    'Legal Holiday Falling on Rest Day_1st 8hrs',
                                    'Legal Holiday Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential Regular Days_1st 8hrs',
                                    'Night Differential Regular Days_Excess of 8hrs',
                                    'Night Differential Falling on Rest Day_1st 8hrs',
                                    'Night Differential Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential falling on Special Holiday',
                                    'Night Differential SH_EX8',
                                    'Night Differential Falling on SPHOL rest day 1st 8 hr',
                                    'Night Differential SH falling on RD_EX8',
                                    'Night Differential on Legal Holidays_1st 8hrs',
                                    'Night Differential on Legal Holidays_Excess of 8hrs',
                                    'Night Differential on Legal Holidays falling on Rest Days'])

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
                                    'RestDay Overtime for the 1st 8hrs',
                                    'Rest Day Overtime in Excess of 8hrs',
                                    'Special Holiday',
                                    'Special Holiday_1st 8hours',
                                    'Special Holiday_Excess of 8hrs',
                                    'Special Holiday Falling on restday 1st 8hrs',
                                    'Special Holiday on restday Excess 8Hrs',
                                    'Legal Holiday',
                                    'Legal Holiday_1st 8hours',
                                    'Legal Holiday_Excess of 8hrs',
                                    'Legal Holiday Falling on Rest Day_1st 8hrs',
                                    'Legal Holiday Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential Regular Days_1st 8hrs',
                                    'Night Differential Regular Days_Excess of 8hrs',
                                    'Night Differential Falling on Rest Day_1st 8hrs',
                                    'Night Differential Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential falling on Special Holiday',
                                    'Night Differential SH_EX8',
                                    'Night Differential Falling on SPHOL rest day 1st 8 hr',
                                    'Night Differential SH falling on RD_EX8',
                                    'Night Differential on Legal Holidays_1st 8hrs',
                                    'Night Differential on Legal Holidays_Excess of 8hrs',
                                    'Night Differential on Legal Holidays falling on Rest Days'])

    for i, row in df.iterrows():
        # Employee inputs
        employee_name = row['Employee Name'].upper()
        employee_code = row['Employee Code'].upper()


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
            work_descript = 'special holiday'

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
                                    'RestDay Overtime for the 1st 8hrs',
                                    'Rest Day Overtime in Excess of 8hrs',
                                    'Special Holiday',
                                    'Special Holiday_1st 8hours',
                                    'Special Holiday_Excess of 8hrs',
                                    'Special Holiday Falling on restday 1st 8hrs',
                                    'Special Holiday on restday Excess 8Hrs',
                                    'Legal Holiday',
                                    'Legal Holiday_1st 8hours',
                                    'Legal Holiday_Excess of 8hrs',
                                    'Legal Holiday Falling on Rest Day_1st 8hrs',
                                    'Legal Holiday Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential Regular Days_1st 8hrs',
                                    'Night Differential Regular Days_Excess of 8hrs',
                                    'Night Differential Falling on Rest Day_1st 8hrs',
                                    'Night Differential Falling on Rest Day_Excess of 8hrs',
                                    'Night Differential falling on Special Holiday',
                                    'Night Differential SH_EX8',
                                    'Night Differential Falling on SPHOL rest day 1st 8 hr',
                                    'Night Differential SH falling on RD_EX8',
                                    'Night Differential on Legal Holidays_1st 8hrs',
                                    'Night Differential on Legal Holidays_Excess of 8hrs',
                                    'Night Differential on Legal Holidays falling on Rest Days'])
    #Employee inputs
    employee_name = request.form.get('employee_name').upper()
    employee_code = request.form.get('employee_code').upper()


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

@app.route('/delete_data', methods=['POST'])
def delete_data():
    csv_file_path = 'dtr.csv'

    if os.path.exists(csv_file_path):
        os.remove(csv_file_path)

    return render_template('index.html')

@app.route('/table', methods=['GET', 'POST'])
def table():
    try:
        df = pd.read_csv('dtr.csv')
        data = df.to_dict('records')
        # add new table with summary information
        summary = pd.DataFrame({'Overtime': [df['Actual Gross Hours Render'].sum()],
                                    'Net Hours Rendered (Time Format)': [df['Net Hours Rendered (Time Format)'].sum()],
                                    'Actual Gross Hours Render': [df['Actual Gross Hours Render'].sum()],
                                    'Hours Rendered': [df['Hours Rendered'].sum()],
                                    'Undertime Hours': [df['Undertime Hours'].sum()],
                                    'Tardiness': [df['Tardiness'].sum()],
                                    'Excess of 8 hours Overtime': [df['Excess of 8 hours Overtime'].sum()],
                                    'Total of 8 hours Overtime': [df['Total of 8 hours Overtime'].sum()],
                                    'RestDay Overtime for the 1st 8hrs': [df['RestDay Overtime for the 1st 8hrs'].sum()],
                                    'Rest Day Overtime in Excess of 8hrs': [df['Rest Day Overtime in Excess of 8hrs'].sum()],
                                    'Special Holiday': [df['Special Holiday'].sum()],
                                    'Special Holiday_1st 8hours': [df['Special Holiday_1st 8hours'].sum()],
                                    'Special Holiday_Excess of 8hrs': [df['Special Holiday_Excess of 8hrs'].sum()],
                                    'Special Holiday Falling on restday 1st 8hrs': [df['Special Holiday Falling on restday 1st 8hrs'].sum()],
                                    'Special Holiday on restday Excess 8Hrs': [df['Special Holiday on restday Excess 8Hrs'].sum()],
                                    'Legal Holiday': [df['Legal Holiday'].sum()],
                                    'Legal Holiday_1st 8hours': [df['Legal Holiday_1st 8hours'].sum()],
                                    'Legal Holiday_Excess of 8hrs': [df['Legal Holiday_Excess of 8hrs'].sum()],
                                    'Legal Holiday Falling on Rest Day_1st 8hrs': [df['Legal Holiday Falling on Rest Day_1st 8hrs'].sum()],
                                    'Legal Holiday Falling on Rest Day_Excess of 8hrs': [df['Legal Holiday Falling on Rest Day_Excess of 8hrs'].sum()],
                                    'Night Differential Regular Days_1st 8hrs': [df['Night Differential Regular Days_1st 8hrs'].sum()],
                                    'Night Differential Regular Days_Excess of 8hrs': [df['Night Differential Regular Days_Excess of 8hrs'].sum()],
                                    'Night Differential Falling on Rest Day_1st 8hrs': [df['Night Differential Falling on Rest Day_1st 8hrs'].sum()],
                                    'Night Differential Falling on Rest Day_Excess of 8hrs': [df['Night Differential Falling on Rest Day_Excess of 8hrs'].sum()] })

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

@app.route('/save_all_rows', methods=['POST'])
def save_all_rows():
    df = pd.read_csv('dtr.csv')
    for index, row in df.iterrows():
        status = request.form.get(f'status_{index}')
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
    app.run(debug=True ,use_reloader=True, port=port)
    # webview.start()
