import json
import os
import webbrowser
from flask import Flask, render_template, request, redirect, send_file, jsonify, url_for
import pandas as pd
from datetime import datetime, timedelta, time
import threading
import platform
import sys
import tempfile
from openpyxl.reader.excel import load_workbook
import traceback

base_dir = '.'
if hasattr(sys, '_MEIPASS'):
    base_dir = os.path.join(sys._MEIPASS)

port = int(os.environ.get('PORT', 9000))

app = Flask(__name__, template_folder=os.path.join(base_dir, 'template'))
app.secret_key = 'topserve_dtr_generator'


def open_browser(url):
    # Open the web page in the default browser
    webbrowser.open(url)


@app.route('/')
def index():
    try:
        dtr = pd.read_csv('dtr.csv')
    except FileNotFoundError:
        dtr = pd.DataFrame(columns=['Status',
                                    'Overtime',
                                    'Employee Code',
                                    'Employee Name',
                                    'Position',
                                    'Cost Center',
                                    'Date',
                                    'Day',
                                    'Working Day',
                                    'Work Description',
                                    'Secondary Description',
                                    'Time In',
                                    'Time Out',
                                    'Actual Time In',
                                    'Actual Time Out',
                                    'Net Hours Rendered (Time Format)',
                                    'Actual Gross Hours Render',
                                    'Hours Rendered',
                                    'Undertime Hours',
                                    'Tardiness',
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
    if y >= .50:
        y = .50
        estimate = x + y

    elif y <= .49:
        y = .0
        estimate = x + y

    if estimate < 1:
        estimate = 0

    return estimate

def adjust_time(input_time):
    time_format = "%H:%M"
    input_datetime = datetime.strptime(input_time, time_format)
    minutes = input_datetime.minute

    if minutes < 30:
        adjusted_datetime = input_datetime.replace(minute=30)
        adjusted_time = adjusted_datetime.strftime(time_format)
    else:
        next_hour = (input_datetime + timedelta(hours=1)).replace(minute=0)
        adjusted_time = next_hour.strftime(time_format)

    return adjusted_time

def halfday(time_str, adjustment_hours):
    # Convert the time string to a datetime object
    time_obj = datetime.strptime(time_str, "%H:%M:%S")

    # Add or subtract the adjustment hours
    adjusted_time = time_obj + timedelta(hours=adjustment_hours)

    # Convert the adjusted time back to a string format
    adjusted_time_str = adjusted_time.strftime("%H:%M:%S")

    return adjusted_time_str

def calculate_timeanddate(employee_name, employee_code, position, location, cost_center, day_of_week, date_transact1, date_obj, work_descript, second_descript, time_in, time_out, time_in1, time_out1, actual_time_in, actual_time_out, actual_time_in1, actual_time_out1):

    #TOTAL OF SCHEDULED TIME IN AND TIME OUT FOR LATE AND UNDERTIME HINDI PARA SA COMPUTATIONG NG OVERTIME
    datetime_timein0 = datetime.combine(date_obj.date(), time_in1)
    if time_in1 > time_out1:
        add_day = timedelta(days=1)
        new_day = date_obj.date() + add_day
        datetime_timeout0 = datetime.combine(new_day, time_out1)

    else:
        datetime_timeout0 = datetime.combine(date_obj.date(), time_out1)

    total_datetime_in_out0 = datetime_timeout0 - datetime_timein0
    total_datetime_in_out_int0 = timedelta_to_decimal(total_datetime_in_out0)



    #FOR LATE AND UNDERTIME
    actual_datetime_timein0 = datetime.combine(date_obj.date(), actual_time_in1)

    if actual_time_in1 > actual_time_out1:
        actual_add_day = timedelta(days=1)
        actual_new_day = date_obj.date() + actual_add_day
        actual_datetime_timeout0 = datetime.combine(actual_new_day, actual_time_out1)

    else:
        actual_datetime_timeout0 = datetime.combine(date_obj.date(), actual_time_out1)

    total_actual_datetime_in_out0 = actual_datetime_timeout0 - actual_datetime_timein0
    total_actual_datetime_in_out_int0 = timedelta_to_decimal(total_actual_datetime_in_out0)

    workingday = 0
    tardiness_str = 0
    undertime_check_try = 0
    if second_descript == "REGULAR DAY":

        check_halfday = actual_datetime_timein0 - datetime_timein0
        check_halfday = timedelta_to_decimal(check_halfday)


        if check_halfday > 2:
            time_in1 = str(time_in1)
            time_in1 = halfday(time_in1, 5)
            time_in1 = datetime.strptime(str(time_in1), "%H:%M:%S").time()
            datetime_timein0 = datetime.combine(date_obj.date(), time_in1)
            workingday = .5

        else:
            workingday = 1

        print("Working Day:", workingday)

    if second_descript == "REGULAR DAY":

        # TARDINESS
        expected = time(0, 0, 0)
        expected_timedelta = timedelta(hours=expected.hour, minutes=expected.minute, seconds=expected.second)

        if actual_datetime_timein0 > datetime_timein0:
            tardiness = actual_datetime_timein0 - datetime_timein0
            tardiness_str = timedelta_to_decimal(tardiness)
            tardiness_str = "{:.2f}".format(tardiness_str)
            tardiness_str = float(tardiness_str)

        else:
            tardiness = expected_timedelta
            tardiness_str = timedelta_to_decimal(tardiness)
            tardiness_str = "{:.2f}".format(tardiness_str)
            tardiness_str = float(tardiness_str)

        #UNDERTIME HOURS

        if second_descript == 'REGULAR' and work_descript == 'REGULAR DAY':
            if actual_datetime_timeout0 < datetime_timeout0:
                undertime_check_try = datetime_timeout0 - actual_datetime_timeout0
                undertime_check_try = timedelta_to_decimal(undertime_check_try)
                undertime_check_try = "{:.2f}".format(undertime_check_try)
                undertime_check_try = float(undertime_check_try)
            else:
                undertime_check_try = 0

        else:
            undertime_check_try = 0

        tardiness_str = tardiness_str + undertime_check_try


    #==========================================================================================================================================
    #TOTAL OF SCHEDULED TIME IN AND TIME OUT
    datetime_timein = datetime.combine(date_obj.date(), time_in1)
    if time_in1 > time_out1:
        add_day = timedelta(days=1)
        new_day = date_obj.date() + add_day
        datetime_timeout = datetime.combine(new_day, time_out1)

    else:
        datetime_timeout = datetime.combine(date_obj.date(), time_out1)

    # ITO YUNG TOTAL BOSS
    total_datetime_in_out = datetime_timeout - datetime_timein
    total_datetime_in_out_int = timedelta_to_decimal(total_datetime_in_out)


    # =====================================================================
    # TOTAL OF ACTUAL TIME IN AND TIME OUT FOR COMPUTATION NG OVERTIME

    # Extract the hours and minutes from the input time
    if actual_time_in1 < time_in1:
        in_hours = actual_time_in1.hour
        in_minutes = actual_time_in1.minute

        # Convert the time to a datetime object with the current date
        current_date = datetime.now().date()
        actual_time_in1_datetime = datetime.combine(current_date, actual_time_in1)

        # Check if minutes is less than 30
        if in_minutes < 30 and in_minutes >= 1:
            # Set the minutes to 30
            actual_time_in1_datetime += timedelta(minutes=30 - in_minutes)
        elif in_minutes > 30 and in_minutes <= 59:
            # Set the minutes to 0 and add 1 hour
            actual_time_in1_datetime += timedelta(minutes=60 - in_minutes)

        # Extract the time from the datetime object
        actual_time_in1 = actual_time_in1_datetime.time()

    #===============================================================
    # Extract the hours and minutes from the input time
    out_hours = actual_time_out1.hour
    out_minutes = actual_time_out1.minute

    # Convert the time to a datetime object with the current date
    current_date = datetime.now().date()
    actual_time_out1_datetime = datetime.combine(current_date, actual_time_out1)

    # Check if minutes is less than 30
    # Check if minutes is less than 30
    if out_minutes >= 30:
        # Set the minutes to 30
        actual_time_out1_datetime += timedelta(minutes=30 - out_minutes)
    else:
        # Set the minutes to 0 and add 1 hour
        actual_time_out1_datetime += timedelta(minutes=-out_minutes)

    # Extract the time from the datetime object
    actual_time_out1 = actual_time_out1_datetime.time()

    # ===============================================================
    actual_datetime_timein = datetime.combine(date_obj.date(), actual_time_in1)

    if actual_time_in1 > actual_time_out1:
        actual_add_day = timedelta(days=1)
        actual_new_day = date_obj.date() + actual_add_day
        actual_datetime_timeout = datetime.combine(actual_new_day, actual_time_out1)

    else:
        actual_datetime_timeout = datetime.combine(date_obj.date(), actual_time_out1)

    if actual_datetime_timeout > datetime_timeout:
        diff = actual_datetime_timeout - datetime_timeout
        diff = timedelta_to_decimal(diff)
        if diff < 1:
            actual_datetime_timeout = datetime_timeout

    print(f"IN: {actual_time_in1} of {employee_name}")
    print(f"OUT: {actual_time_out1} of {employee_name}")

    # ITO YUNG TOTAL BOSS
    total_actual_datetime_in_out = actual_datetime_timeout - actual_datetime_timein
    total_actual_datetime_in_out_int = timedelta_to_decimal(total_actual_datetime_in_out)

    #=====================================================================
    #NON CHARGEABLE BREAK

    if total_actual_datetime_in_out_int >= 5:
        non_chargeable_break = time(1, 0, 0)
        non_chargeable_break_timedelta = timedelta(hours=non_chargeable_break.hour, minutes=non_chargeable_break.minute,
                                                   seconds=non_chargeable_break.second)


    else:
        non_chargeable_break = time(0, 0, 0)
        non_chargeable_break_timedelta = timedelta(hours=non_chargeable_break.hour, minutes=non_chargeable_break.minute,
                                                   seconds=non_chargeable_break.second)

    #=====================================================================
    #NET HOURS RENDERED
    if second_descript == 'REGULAR DAY' and work_descript == 'REGULAR':
        net_hours_rendered = total_datetime_in_out - non_chargeable_break_timedelta
        net_hours_rendered_str = timedelta_to_decimal(net_hours_rendered)

    else:
        net_hours_rendered_str = 0

    #=====================================================================
    # ACTUAL GROSS RENDERED

    actual_render = 0
    if total_actual_datetime_in_out_int > 0:
        actual_render = total_actual_datetime_in_out - non_chargeable_break_timedelta
        actual_render = timedelta_to_decimal(actual_render)
        actual_render = "{:.2f}".format(actual_render)
        actual_render = float(actual_render)


    elif work_descript == 'LEGAL HOLIDAY' and total_actual_datetime_in_out_int == 0:
        actual_render = 0

    elif work_descript == 'SPECIAL HOLIDAY' and total_actual_datetime_in_out_int == 0:
        actual_render = 0

    print("Total actual shift", total_actual_datetime_in_out)

    #=====================================================================
    # HOURS RENDERED

    hour_rendered = total_actual_datetime_in_out - non_chargeable_break_timedelta
    hour_rendered_str = timedelta_to_decimal(hour_rendered)
    hour_rendered_str = int(hour_rendered_str)


    if work_descript == 'LEGAL HOLIDAY' or work_descript == 'SPECIAL HOLIDAY' and total_actual_datetime_in_out_int == 0 and total_datetime_in_out_int == 0:
        hour_rendered_str = 0


    if hour_rendered_str > 8:
        hour_rendered_str = 8


    #=====================================================================
    # OVERTIME

    # NGHT DIFFERENCE LODS
    night_diff_start = '22:00'
    night_diff_end = '06:00'

    datetime_start_night = datetime.strptime(night_diff_start, "%H:%M").time()
    datetime_start_night = datetime.combine(date_obj, datetime_start_night)
    datetime_end_night_before = datetime.strptime(night_diff_end, "%H:%M").time()
    datetime_end_night_before1 = datetime.combine(date_obj, datetime_end_night_before)
    one_day = timedelta(days=1)
    new_date = date_obj + one_day
    date_time_end_night = datetime.combine(new_date, datetime_end_night_before)


    # VARIABLES TO LODS
    overtime = 0

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

    ND_regular_days_1st_8_hrsnotOT = 0

    ND_regular_days_1st_8_hrs = 0
    ND_regular_days_excess_8_hrs = 0
    ND_regular_days_excess_8_hrs1 = 0

    ND_regular_days_RD_1st_8_hrs = 0
    ND_regular_days_RD_excess_8_hrs = 0

    ND_legal_holiday_1st_8_hrs = 0
    ND_legal_holiday_excess_8_hrs = 0

    ND_legal_holiday_RD_1st_8_hrs = 0
    ND_legal_holiday_RD_excess_8_hrs = 0

    ND_special_holiday_1st_8_hrs = 0
    ND_special_holiday_excess_8_hrs = 0

    ND_special_holiday_RD_1st_8_hrs = 0
    ND_special_holiday_RD_excess_8_hrs = 0



    if work_descript == 'REGULAR' and second_descript == 'REGULAR DAY':
        print("Regular and Regular Day")
        # These conditions are for Night Difference within the regular hours
        if actual_datetime_timein > datetime_start_night and actual_datetime_timeout > date_time_end_night:
            ND_regular_days_1st_8_hrs = date_time_end_night - actual_datetime_timein
            ND_regular_days_1st_8_hrs = timedelta_to_decimal(ND_regular_days_1st_8_hrs)
            print("Pumasok sa Night Difference condition 1")

        elif actual_datetime_timein < datetime_start_night and actual_datetime_timeout > datetime_start_night and datetime_timein < datetime_start_night and datetime_timeout > datetime_start_night:
            ND_regular_days_1st_8_hrs = actual_datetime_timeout - datetime_start_night
            ND_regular_days_1st_8_hrs = timedelta_to_decimal(ND_regular_days_1st_8_hrs)
            print("Pumasok sa Night Difference condition 2")

        elif actual_datetime_timein < datetime_end_night_before1 and datetime_timein < datetime_end_night_before1:
            ND_regular_days_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
            ND_regular_days_1st_8_hrs = timedelta_to_decimal(ND_regular_days_1st_8_hrs)
            print("Pumasok sa Night Difference condition 3")

        #Para sa overtime na pinutol sa regular hours
        if actual_datetime_timein > datetime_timeout or actual_datetime_timeout < datetime_timein:
            print("Pumasok sa overtime na pinutol sa regular hours")

            undertime_check_try = 0
            workingday = 0
            tardiness_str = 0
            overtime = total_actual_datetime_in_out_int


            # These conditions are for Night Difference within the Overtime
            if actual_datetime_timeout > datetime_start_night:
                Night_diff = actual_datetime_timeout - datetime_start_night
                Night_diff = timedelta_to_decimal(Night_diff)
                ND_regular_days_excess_8_hrs1 = overtime - Night_diff




        else:
            print("Pumasok sa overtime")
            # This condition is for calculating the overtime
            if actual_datetime_timeout > datetime_timeout or actual_datetime_timein < datetime_timein:
                print("Pumasok sa overtime")
                early_in = 0
                if actual_datetime_timein < datetime_timein:
                    print("May early in")
                    early_in = datetime_timein - actual_datetime_timein
                    early_in = timedelta_to_decimal(early_in)
                overtime = actual_datetime_timeout - datetime_timeout
                overtime = timedelta_to_decimal(overtime)
                overtime = overtime + early_in


                # These conditions are for Night Difference within the Overtime
                if actual_datetime_timeout > datetime_start_night:
                    print("Night Difference within the Regular Overtime")
                    Night_diff = actual_datetime_timeout - datetime_start_night
                    ND_regular_days_excess_8_hrs = timedelta_to_decimal(Night_diff)

                elif actual_datetime_timein < datetime_end_night_before1:
                    print("Night Difference within the Regular Overtime early in")
                    Night_diff = datetime_end_night_before1 - actual_datetime_timein
                    Night_diff = timedelta_to_decimal(Night_diff)
                    ND_regular_days_excess_8_hrs1 = Night_diff



            # This line of code is for subtracting the hours outside the Night difference.
            # For future coder sorry sakit na ng ulo ko yan na naisip kong solution sa problem hahahaha
        if actual_datetime_timeout > date_time_end_night:
            cancel_notND = actual_datetime_timeout - date_time_end_night
            cancel_notND = timedelta_to_decimal(cancel_notND)
            ND_regular_days_excess_8_hrs = ND_regular_days_excess_8_hrs - cancel_notND

        ND_regular_days_excess_8_hrs = ND_regular_days_excess_8_hrs + ND_regular_days_excess_8_hrs1
        overtime = hour_estimate(overtime)

    elif second_descript == 'REST DAY' and work_descript == 'REST DAY':
        print("REST DAY and REST DAY")
        if actual_render > 8:
            restday_ot_excess_8 = actual_render - 8
            restday_ot_first_8 = actual_render - restday_ot_excess_8
            restday_ot_excess_8 = "{:.2f}".format(restday_ot_excess_8)
            restday_ot_excess_8 = float(restday_ot_excess_8)
            overtime = 0

            if actual_datetime_timeout > datetime_start_night:

                start_shift = max(datetime_start_night, actual_datetime_timein)
                end_shift = min(date_time_end_night, actual_datetime_timein + timedelta(hours=9))
                Night_diff = max(timedelta(), end_shift - start_shift)
                excess_ND = actual_datetime_timeout - datetime_start_night
                Night_diff = timedelta_to_decimal(Night_diff)
                excess_ND = timedelta_to_decimal(excess_ND)
                excess_diff = excess_ND - Night_diff

                ND_regular_days_RD_excess_8_hrs = excess_diff
                ND_regular_days_RD_1st_8_hrs = Night_diff

            if actual_datetime_timein < datetime_end_night_before1:
                ND_regular_days_RD_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                ND_regular_days_RD_1st_8_hrs = timedelta_to_decimal(ND_regular_days_RD_1st_8_hrs)

        elif actual_render <= 8:
            restday_ot_first_8 = actual_render
            overtime = 0
            restday_ot_first_8 = hour_estimate(restday_ot_first_8)

            if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout < datetime_end_night_before1:
                Night_diff = actual_render
                ND_overtime = Night_diff
                ND_regular_days_RD_1st_8_hrs = ND_overtime

            if actual_datetime_timeout > datetime_start_night:
                Night_diff = actual_datetime_timeout - datetime_start_night
                Night_diff = timedelta_to_decimal(Night_diff)
                ND_overtime = Night_diff
                ND_regular_days_RD_1st_8_hrs = ND_overtime

            if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout > datetime_end_night_before1:
                ND_regular_days_RD_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                ND_regular_days_RD_1st_8_hrs = timedelta_to_decimal(ND_regular_days_RD_1st_8_hrs)


        restday_ot_first_8 = hour_estimate(restday_ot_first_8)
        restday_ot_excess_8 = hour_estimate(restday_ot_excess_8)


    #IF DI PUMASOK YUNG EMPLOYEE
    elif second_descript == 'LEGAL HOLIDAY' and work_descript == 'REGULAR' and total_actual_datetime_in_out_int == 0:
        overtime = 0
        legal_holiday = 8

    #IF PUMASOK YUNG EMPLOYEE DITO MAGCOCOMPUTE
    elif second_descript == 'LEGAL HOLIDAY' and total_actual_datetime_in_out_int > 0 and total_datetime_in_out_int > 0:
        if work_descript == 'REGULAR':
            print("Regular and Legal Holiday")
            if actual_render >= 8:
                legal_holiday_excess_8 = actual_render - 8
                legal_holiday_1st_8 = actual_render - legal_holiday_excess_8
                legal_holiday_excess_8 = "{:.2f}".format(legal_holiday_excess_8)
                legal_holiday_excess_8 = float(legal_holiday_excess_8)

                #These conditions are for Night Difference
                if actual_datetime_timeout > datetime_start_night:

                    start_shift = max(datetime_start_night, actual_datetime_timein)
                    end_shift = min(date_time_end_night, actual_datetime_timein + timedelta(hours=9))
                    Night_diff = max(timedelta(), end_shift - start_shift)
                    excess_ND = actual_datetime_timeout - datetime_start_night
                    Night_diff = timedelta_to_decimal(Night_diff)
                    excess_ND = timedelta_to_decimal(excess_ND)
                    excess_diff = excess_ND - Night_diff

                    if ND_legal_holiday_1st_8_hrs >= 8:
                        ND_legal_holiday_excess_8_hrs= excess_diff
                    ND_legal_holiday_1st_8_hrs = Night_diff


                if actual_datetime_timein < datetime_end_night_before1:
                    ND_legal_holiday_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                    ND_legal_holiday_1st_8_hrs = timedelta_to_decimal(ND_legal_holiday_1st_8_hrs)




            elif actual_render <= 8:
                legal_holiday_1st_8 = actual_render
                legal_holiday_1st_8 = "{:.2f}".format(legal_holiday_1st_8)
                legal_holiday_1st_8 = float(legal_holiday_1st_8)

                # These conditions are for Night Difference
                if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout < datetime_end_night_before1:
                    Night_diff = legal_holiday_1st_8
                    ND_overtime = Night_diff
                    ND_legal_holiday_1st_8_hrs = ND_overtime

                if actual_datetime_timeout > datetime_start_night and actual_datetime_timein < datetime_start_night:
                    Night_diff = actual_datetime_timeout - datetime_start_night
                    Night_diff = timedelta_to_decimal(Night_diff)
                    ND_overtime = Night_diff
                    ND_legal_holiday_1st_8_hrs = ND_overtime

                if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout > datetime_end_night_before1:
                    ND_legal_holiday_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                    ND_legal_holiday_1st_8_hrs = timedelta_to_decimal(ND_legal_holiday_1st_8_hrs)



        elif work_descript == 'REST DAY':
            if actual_render > 8:
                legal_holiday_rd_excess_8 = actual_render - 8
                legal_holiday_rd_1st_8 = actual_render - legal_holiday_rd_excess_8
                legal_holiday_rd_excess_8 = "{:.2f}".format(legal_holiday_rd_excess_8)
                legal_holiday_rd_excess_8 = float(legal_holiday_rd_excess_8)

                # These conditions are for Night Difference
                if actual_datetime_timeout > datetime_start_night:

                    start_shift = max(datetime_start_night, actual_datetime_timein)
                    end_shift = min(date_time_end_night, actual_datetime_timein + timedelta(hours=9))
                    Night_diff = max(timedelta(), end_shift - start_shift)
                    excess_ND = actual_datetime_timeout - datetime_start_night
                    Night_diff = timedelta_to_decimal(Night_diff)
                    excess_ND = timedelta_to_decimal(excess_ND)
                    excess_diff = excess_ND - Night_diff

                    if ND_legal_holiday_RD_1st_8_hrs >= 8:
                        ND_legal_holiday_RD_excess_8_hrs = excess_diff
                    ND_legal_holiday_RD_1st_8_hrs = Night_diff

                if actual_datetime_timein < datetime_end_night_before1:
                    ND_legal_holiday_RD_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                    ND_legal_holiday_RD_1st_8_hrs = timedelta_to_decimal(ND_legal_holiday_RD_1st_8_hrs)

            elif actual_render <= 8:


                legal_holiday_rd_1st_8 = actual_render
                legal_holiday_rd_1st_8 = "{:.2f}".format(legal_holiday_rd_1st_8)
                legal_holiday_rd_1st_8 = float(legal_holiday_rd_1st_8)

                # These conditions are for Night Difference
                if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout < datetime_end_night_before1:
                    Night_diff = actual_render
                    ND_overtime = Night_diff
                    ND_legal_holiday_RD_1st_8_hrs = ND_overtime

                if actual_datetime_timeout > datetime_start_night and actual_datetime_timein < datetime_start_night:
                    Night_diff = actual_datetime_timeout - datetime_start_night
                    Night_diff = timedelta_to_decimal(Night_diff)
                    ND_overtime = Night_diff
                    ND_legal_holiday_RD_1st_8_hrs = ND_overtime

                if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout > datetime_end_night_before1:
                    ND_legal_holiday_RD_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                    ND_legal_holiday_RD_1st_8_hrs = timedelta_to_decimal(ND_legal_holiday_RD_1st_8_hrs)


        legal_holiday_1st_8 = hour_estimate(legal_holiday_1st_8)
        legal_holiday_excess_8 = hour_estimate(legal_holiday_excess_8)
        legal_holiday_rd_1st_8 = hour_estimate(legal_holiday_rd_1st_8)
        legal_holiday_rd_excess_8 = hour_estimate(legal_holiday_rd_excess_8)

    elif second_descript == 'SPECIAL HOLIDAY' and total_actual_datetime_in_out_int > 0 and total_datetime_in_out_int > 0:
        if work_descript == 'REGULAR':
            print("Regular and special Holiday")
            if actual_render > 8:
                special_holiday_excess_8 = actual_render - 8
                special_holiday_1st_8 = actual_render - special_holiday_excess_8
                special_holiday_excess_8 = "{:.2f}".format(special_holiday_excess_8)
                special_holiday_excess_8 = float(special_holiday_excess_8)

                # These conditions are for Night Difference
                if actual_datetime_timeout > datetime_start_night:
                    start_shift = max(datetime_start_night, actual_datetime_timein)
                    end_shift = min(date_time_end_night, actual_datetime_timein + timedelta(hours=9))
                    Night_diff = max(timedelta(), end_shift - start_shift)
                    excess_ND = actual_datetime_timeout - datetime_start_night
                    Night_diff = timedelta_to_decimal(Night_diff)
                    excess_ND = timedelta_to_decimal(excess_ND)
                    excess_diff = excess_ND - Night_diff

                    if ND_special_holiday_1st_8_hrs >= 8:
                        ND_special_holiday_excess_8_hrs = excess_diff
                    ND_special_holiday_1st_8_hrs = Night_diff


                if actual_datetime_timein < datetime_end_night_before1:
                    ND_special_holiday_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                    ND_special_holiday_1st_8_hrs = timedelta_to_decimal(ND_special_holiday_1st_8_hrs)


            elif actual_render <= 8:
                special_holiday_1st_8 = actual_render
                special_holiday_1st_8 = "{:.2f}".format(special_holiday_1st_8)
                special_holiday_1st_8 = float(special_holiday_1st_8)

                # These conditions are for Night Difference
                if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout < datetime_end_night_before1:
                    Night_diff = actual_render
                    ND_overtime = Night_diff
                    ND_special_holiday_1st_8_hrs = ND_overtime

                if actual_datetime_timeout > datetime_start_night:
                    Night_diff = actual_datetime_timeout - datetime_start_night
                    Night_diff = timedelta_to_decimal(Night_diff)
                    ND_overtime = Night_diff
                    ND_special_holiday_1st_8_hrs = ND_overtime

                if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout > datetime_end_night_before1:
                    ND_special_holiday_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                    ND_special_holiday_1st_8_hrs = timedelta_to_decimal(ND_special_holiday_1st_8_hrs)


        elif work_descript == 'REST DAY':
            print("Restday and special Holiday")
            if actual_render > 8:
                special_holiday_rd_excess_8 = actual_render - 8
                special_holiday_rd_1st_8 = actual_render - special_holiday_rd_excess_8
                special_holiday_rd_excess_8 = "{:.2f}".format(special_holiday_rd_excess_8)
                special_holiday_rd_excess_8 = float(special_holiday_rd_excess_8)

                # These conditions are for Night Difference
                if actual_datetime_timeout > datetime_start_night:
                    start_shift = max(datetime_start_night, actual_datetime_timein)
                    end_shift = min(date_time_end_night, actual_datetime_timein + timedelta(hours=9))
                    Night_diff = max(timedelta(), end_shift - start_shift)
                    excess_ND = actual_datetime_timeout - datetime_start_night
                    Night_diff = timedelta_to_decimal(Night_diff)
                    excess_ND = timedelta_to_decimal(excess_ND)
                    excess_diff = excess_ND - Night_diff

                    if ND_special_holiday_RD_1st_8_hrs >= 8:
                        ND_special_holiday_RD_excess_8_hrs = excess_diff
                    ND_special_holiday_RD_1st_8_hrs = Night_diff

                if actual_datetime_timein < datetime_end_night_before1:
                    ND_special_holiday_RD_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                    ND_special_holiday_RD_1st_8_hrs = timedelta_to_decimal(ND_special_holiday_RD_1st_8_hrs)


            elif actual_render <= 8:
                special_holiday_rd_1st_8 = actual_render
                special_holiday_rd_1st_8 = "{:.2f}".format(special_holiday_rd_1st_8)
                special_holiday_rd_1st_8 = float(special_holiday_rd_1st_8)

                # These conditions are for Night Difference
                if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout < datetime_end_night_before1:
                    Night_diff = actual_render
                    ND_overtime = Night_diff
                    ND_special_holiday_RD_1st_8_hrs = ND_overtime

                if actual_datetime_timeout > datetime_start_night:
                    Night_diff = actual_datetime_timeout - datetime_start_night
                    Night_diff = timedelta_to_decimal(Night_diff)
                    ND_overtime = Night_diff
                    ND_special_holiday_RD_1st_8_hrs = ND_overtime

                if actual_datetime_timein < datetime_end_night_before1 and actual_datetime_timeout > datetime_end_night_before1:
                    ND_special_holiday_RD_1st_8_hrs = datetime_end_night_before1 - actual_datetime_timein
                    ND_special_holiday_RD_1st_8_hrs = timedelta_to_decimal(ND_special_holiday_RD_1st_8_hrs)


            special_holiday_excess_8 = hour_estimate(special_holiday_excess_8)
            special_holiday_1st_8 = hour_estimate(special_holiday_1st_8)
            special_holiday_rd_excess_8 = hour_estimate(special_holiday_rd_excess_8)
            special_holiday_rd_1st_8 = hour_estimate(special_holiday_rd_1st_8)

    ND_regular_days_1st_8_hrs = hour_estimate(ND_regular_days_1st_8_hrs)
    ND_regular_days_excess_8_hrs = hour_estimate(ND_regular_days_excess_8_hrs)
    ND_regular_days_RD_1st_8_hrs = hour_estimate(ND_regular_days_RD_1st_8_hrs)
    ND_regular_days_RD_excess_8_hrs = hour_estimate(ND_regular_days_RD_excess_8_hrs)
    ND_legal_holiday_1st_8_hrs = hour_estimate(ND_legal_holiday_1st_8_hrs)
    ND_legal_holiday_excess_8_hrs = hour_estimate(ND_legal_holiday_excess_8_hrs)
    ND_legal_holiday_RD_1st_8_hrs = hour_estimate(ND_legal_holiday_RD_1st_8_hrs)
    ND_legal_holiday_RD_excess_8_hrs = hour_estimate(ND_legal_holiday_RD_excess_8_hrs)
    ND_special_holiday_1st_8_hrs = hour_estimate(ND_special_holiday_1st_8_hrs)
    ND_special_holiday_excess_8_hrs = hour_estimate(ND_special_holiday_excess_8_hrs)
    ND_special_holiday_RD_1st_8_hrs = hour_estimate(ND_special_holiday_RD_1st_8_hrs)
    ND_special_holiday_RD_excess_8_hrs = hour_estimate(ND_special_holiday_RD_excess_8_hrs)

    print("Overtime", overtime)
    print("")
    print("========================================================")
    print("")




    #=====================================================================
    # Approval
    status_OT = 'No OT'
    if overtime >= 1:
        status_OT = 'Not Approved'
    elif work_descript == 'REGULAR DAY' and overtime == 0:
        status_OT = 'No OT'
    elif work_descript == 'LEGAL HOLIDAY' or work_descript == 'SPECIAL HOLIDAY' and overtime == 0:
        status_OT = 'Holiday'

    # Encoding of inputs to the dataframe
    dtr_new = pd.DataFrame({'Status': [status_OT],
                            'Overtime': [overtime],
                            'Employee Code': [employee_code],
                            'Employee Name': [employee_name],
                            'Position': [position],
                            'Location': [location],
                            'Cost Center': [cost_center],
                            'Date': [date_transact1],
                            'Day': [day_of_week],
                            'Working Day': [workingday],
                            'Work Description': [work_descript],
                            'Secondary Description': [second_descript],
                            'Time In': [time_in],
                            'Time Out': [time_out],
                            'Actual Time In': [actual_time_in],
                            'Actual Time Out': [actual_time_out],
                            'Net Hours Rendered (Time Format)': [net_hours_rendered_str],
                            'Actual Gross Hours Render': [actual_render],
                            'Hours Rendered': [hour_rendered_str],
                            'Undertime Hours': [undertime_check_try],
                            'Tardiness': [tardiness_str],
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
                            'Night Differential Falling on Rest Day_Excess of 8hrs': [ND_regular_days_RD_excess_8_hrs],
                            'Night Differential falling on Special Holiday': [ND_special_holiday_1st_8_hrs],
                            'Night Differential SH_EX8': [ND_special_holiday_excess_8_hrs],
                            'Night Differential Falling on SPHOL rest day 1st 8 hr': [ND_special_holiday_RD_1st_8_hrs],
                            'Night Differential SH falling on RD_EX8': [ND_special_holiday_RD_excess_8_hrs],
                            'Night Differential on Legal Holidays_1st 8hrs': [ND_legal_holiday_1st_8_hrs],
                            'Night Differential on Legal Holidays_Excess of 8hrs': [ND_legal_holiday_excess_8_hrs],
                            'Night Differential on Legal Holidays falling on Rest Days': [ND_legal_holiday_RD_1st_8_hrs]
                            })

    try:
        dtr = pd.read_csv('dtr.csv')
    except FileNotFoundError:
        dtr = pd.DataFrame(columns=['Status',
                                    'Overtime',
                                    'Employee Code',
                                    'Employee Name',
                                    'Position',
                                    'Location',
                                    'Cost Center',
                                    'Date',
                                    'Day',
                                    'Working Day',
                                    'Work Description',
                                    'Secondary Description',
                                    'Time In',
                                    'Time Out',
                                    'Actual Time In',
                                    'Actual Time Out',
                                    'Net Hours Rendered (Time Format)',
                                    'Actual Gross Hours Render',
                                    'Hours Rendered',
                                    'Undertime Hours',
                                    'Tardiness',
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
    try:
        # Get the uploaded file
        uploaded_file = request.files['file']
        # Read the Excel file into a pandas dataframe
        df = pd.read_excel(uploaded_file, header=[0], engine='openpyxl')

        dtr = pd.DataFrame(columns=['Status',
                                    'Overtime',
                                    'Employee Code',
                                    'Employee Name',
                                    'Position',
                                    'Location',
                                    'Cost Center',
                                    'Date',
                                    'Day',
                                    'Working Day',
                                    'Work Description',
                                    'Secondary Description',
                                    'Time In',
                                    'Time Out',
                                    'Actual Time In',
                                    'Actual Time Out',
                                    'Net Hours Rendered (Time Format)',
                                    'Actual Gross Hours Render',
                                    'Hours Rendered',
                                    'Undertime Hours',
                                    'Tardiness',
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
            employee_name = row['Employee Name']
            employee_code = row['Employee Code']

            if not isinstance(employee_name, str):  # Use isinstance() to check if employee_name is not a string
                print("End of Uploading")
                break

            if isinstance(employee_name, str):
                employee_name = employee_name.upper()

            if isinstance(employee_code, str):
                employee_code = employee_code.upper()

            position = row['Position'].upper()
            location = row['Location'].upper()
            cost_center = row['Cost Center'].upper()

            # Date of transaction
            date_obj = ""
            try:
                date_transact = str(row['Date'])
                day_of_week = ""
                date_transact1 = ""
                if date_transact is not None:
                    date_obj = datetime.strptime(date_transact, "%Y-%m-%d %H:%M:%S")
                    date_transact1 = date_obj.date()
                    day_of_week = date_obj.strftime("%A")
                    day_of_week = day_of_week.upper()
                else:
                    date_obj = ""
            except Exception:
                # Handle the exception and pass the error message to the template
                error_message = f"Invalid value {date_obj} in the Date column. Allowed format dd/mm/yyyy."
                traceback.print_exc()  # Print the traceback for debugging purposes
                return render_template('index.html', error_message=error_message)



            # Description inputs
            work_descript = row['Work Description'].upper()
            # if work_descript.upper() == 'REGULAR':
            #     work_descript = 'REGULAR'
            # elif work_descript.upper() == 'REST DAY':
            #     work_descript = 'REST DAY'
            # else:
            #     error_message = f"Invalid value {work_descript} in the Work Description column. Allowed values are regular day, legal holiday, and special holiday."
            #     return render_template('index.html', error_message=error_message)

            second_descript = row['Secondary Description'].upper()

            # Work Hours Inputs
            time_in1 = datetime.strptime(str(row['WS Time In']), "%H:%M:%S").time()
            time_in = time_in1.strftime("%H:%M")
            time_out1 = datetime.strptime(str(row['WS Time Out']), "%H:%M:%S").time()
            time_out = time_out1.strftime("%H:%M")

            # Actual Time in and out
            actual_time_in1 = datetime.strptime(str(row['Actual Time In']), "%H:%M:%S").time()
            actual_time_out1 = datetime.strptime(str(row['Actual Time Out']), "%H:%M:%S").time()
            actual_time_in = actual_time_in1.strftime("%H:%M")
            actual_time_out = actual_time_out1.strftime("%H:%M")

            calculate_timeanddate(employee_name, employee_code, position, location, cost_center, day_of_week, date_transact1, date_obj, work_descript, second_descript,
                                  time_in, time_out, time_in1, time_out1, actual_time_in, actual_time_out,
                                  actual_time_in1, actual_time_out1)




        return render_template('index.html', dtr=dtr)

    except Exception as e:
        # Handle the exception and pass the error message to the template
        error_message = "Error in the uploaded file. Please make sure the file format is correct and no missing values."
        traceback.print_exc()  # Print the traceback for debugging purposes
        return render_template('index.html', error_message=error_message)


@app.route('/submit', methods=['POST'])
def submit():

    #Creating a new Dataframe (this serves as a database na rin)
    dtr = pd.DataFrame(columns=['Status',
                                'Overtime',
                                'Employee Code',
                                'Employee Name',
                                'Position',
                                'Location',
                                'Cost Center',
                                'Date',
                                'Day',
                                'Working Day',
                                'Work Description',
                                'Secondary Description',
                                'Time In',
                                'Time Out',
                                'Actual Time In',
                                'Actual Time Out',
                                'Net Hours Rendered (Time Format)',
                                'Actual Gross Hours Render',
                                'Hours Rendered',
                                'Undertime Hours',
                                'Tardiness',
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
    position = request.form.get('position').upper()
    location = request.form.get('location').upper()
    cost_center = request.form.get('cost_center').upper()



    #Date of transaction
    date_transact = request.form.get('start_date')
    day_of_week = ""
    date_transact1 = ""
    if date_transact is not None:
        date_obj = datetime.strptime(date_transact, "%Y-%m-%d")
        date_transact1 = date_obj.date()
        day_of_week = date_obj.strftime("%A")
        day_of_week = day_of_week.upper()
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
    work_descript = request.form.get('work_descript').upper()
    second_descript = request.form.get('second_descript').upper()

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

    calculate_timeanddate(employee_name, employee_code, position, location, cost_center, day_of_week, date_transact1, date_obj, work_descript, second_descript, time_in, time_out, time_in1, time_out1, actual_time_in, actual_time_out, actual_time_in1, actual_time_out1)


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
        summary = pd.DataFrame({'Overtime': [df['Overtime'].sum()],
                                'Working Day': [df['Working Day'].sum()],
                                    'Net Hours Rendered (Time Format)': [df['Net Hours Rendered (Time Format)'].sum()],
                                    'Actual Gross Hours Render': [df['Actual Gross Hours Render'].sum()],
                                    'Hours Rendered': [df['Hours Rendered'].sum()],
                                    'Undertime Hours': [df['Undertime Hours'].sum()],
                                    'Tardiness': [df['Tardiness'].sum()],
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
                                    'Night Differential Falling on Rest Day_Excess of 8hrs': [df['Night Differential Falling on Rest Day_Excess of 8hrs'].sum()],
                                    'Night Differential falling on Special Holiday': [df['Night Differential falling on Special Holiday'].sum()],
                                    'Night Differential SH_EX8': [df['Night Differential SH_EX8'].sum()],
                                    'Night Differential Falling on SPHOL rest day 1st 8 hr': [df['Night Differential Falling on SPHOL rest day 1st 8 hr'].sum()],
                                    'Night Differential SH falling on RD_EX8': [df['Night Differential SH falling on RD_EX8'].sum()],
                                    'Night Differential on Legal Holidays_1st 8hrs': [df['Night Differential on Legal Holidays_1st 8hrs'].sum()],
                                    'Night Differential on Legal Holidays_Excess of 8hrs': [df['Night Differential on Legal Holidays_Excess of 8hrs'].sum()],
                                    'Night Differential on Legal Holidays falling on Rest Days': [df['Night Differential on Legal Holidays falling on Rest Days'].sum()]
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

@app.route('/save_all_rows', methods=['POST'])
def save_all_rows():
    df = pd.read_csv('dtr.csv')
    for index, row in df.iterrows():
        status = request.form.get(f'status_{index}')
        df.loc[index, 'Status'] = status
    df.to_csv('dtr.csv', index=False)
    return redirect(url_for('table'))

@app.route('/checkcsv')
def check_csv():
    key_error = True
    try:
        df = pd.read_csv('dtr.csv')
        empty = df.empty
    except FileNotFoundError:
        empty = True
    except KeyError:
        empty = False
        key_error = True
    else:
        empty = False
        key_error = False

    response = {
        'empty': empty,
        'keyError': key_error if 'key_error' in locals() else False
    }
    return json.dumps(response)

@app.route('/download')
def download():
    file_name1 = 'DTR_excel.xlsx'
    if os.path.exists(file_name1):
        os.remove(file_name1)

    try:
        df = pd.read_csv('dtr.csv')
    except FileNotFoundError:
        # Handle the exception and pass the error message to the template
        error_message = "No Data to download"
        traceback.print_exc()  # Print the traceback for debugging purposes
        return render_template('index.html', error_message=error_message)

    file_name = 'DTR_Summary.csv'
    if os.path.exists(file_name):
        os.remove(file_name)



    grouped_df = df.groupby(['Employee Name', 'Cost Center'])

    # Step 4: Compute overtime for each employee
    for (name, cost_center), employee_df in grouped_df:
        employee_code = employee_df['Employee Code'].iloc[0]
        location = employee_df['Location'].iloc[0]
        position = employee_df['Position'].iloc[0]
        status = employee_df['Status']

        # Check if the status is not approved and set overtime to 0
        # if 'Not Approved' in status.values:
        #     employee_df.loc[status == 'Not Approved', 'Overtime'] = 0


        total_working_days = employee_df.loc[(employee_df['Secondary Description'] == 'REGULAR DAY') & (employee_df['Work Description'] == 'REGULAR'), 'Working Day'].sum()
        working_hours1 = employee_df.loc[(employee_df['Secondary Description'] == 'REGULAR DAY') & (employee_df['Work Description'] == 'REGULAR'), 'Hours Rendered'].sum()
        tardiness1 = employee_df.loc[(employee_df['Secondary Description'] == 'REGULAR DAY') & (employee_df['Work Description'] == 'REGULAR'), 'Tardiness'].sum()
        RegularDay_Overtime = employee_df['Overtime'].sum()
        RegularDay_Overtime_Excess = 0
        RegularDay_RestDay = 0
        RegularDay_RestDay_Overtime = employee_df['RestDay Overtime for the 1st 8hrs'].sum()
        RegularDay_RestDay_Overtime_Excess = employee_df['Rest Day Overtime in Excess of 8hrs'].sum()
        Specialholiday = employee_df['Special Holiday'].sum()
        Specialholiday_Overtime = employee_df['Special Holiday_1st 8hours'].sum()
        Specialholiday_Overtime_Excess = employee_df['Special Holiday_Excess of 8hrs'].sum()
        Specialholiday_RestDay_Overtime = employee_df['Special Holiday Falling on restday 1st 8hrs'].sum()
        Specialholiday_Restday_Overtime_Excess = employee_df['Special Holiday on restday Excess 8Hrs'].sum()
        legalholiday = employee_df['Legal Holiday'].sum()
        legalholiday_Overtime = employee_df['Legal Holiday_1st 8hours'].sum()
        legalholiday_Overtime_Excess = employee_df['Legal Holiday_Excess of 8hrs'].sum()
        legalholiday_RestDay_Overtime = employee_df['Legal Holiday Falling on Rest Day_1st 8hrs'].sum()
        legalholiday_RestDay_Overtime_Excess = employee_df['Legal Holiday Falling on Rest Day_Excess of 8hrs'].sum()
        RegularDay_Night_8hours = employee_df['Night Differential Regular Days_1st 8hrs'].sum()
        RegularDay_Night_8hours_Excess = employee_df['Night Differential Regular Days_Excess of 8hrs'].sum()
        RegularDay_RestDay_Night_8hours = employee_df['Night Differential Falling on Rest Day_1st 8hrs'].sum()
        RegularDay_RestDay_Night_8hours_Excess = employee_df['Night Differential Falling on Rest Day_Excess of 8hrs'].sum()
        Specialholiday_RestDay_Night_8hours = employee_df['Night Differential Falling on SPHOL rest day 1st 8 hr'].sum()
        Specialholiday_RestDay_Night_8hours_Excess = employee_df['Night Differential SH falling on RD_EX8'].sum()
        legalholiday_rest_Night_diffence = employee_df['Night Differential on Legal Holidays falling on Rest Days'].sum()
        legalholiday_Night_8hours = employee_df['Night Differential on Legal Holidays_1st 8hrs'].sum()
        legalholiday_Night_8hours_Excess = employee_df['Night Differential on Legal Holidays_Excess of 8hrs'].sum()
        Specialholiday_Night_diffence = employee_df['Night Differential falling on Special Holiday'].sum()
        Specialholiday_Night_8hours_Excess = employee_df['Night Differential SH_EX8'].sum()


        new_df = pd.DataFrame({'Employee_Code': [employee_code],
                               'Employee_Name': [name],
                               'Position': [position],
                               'Location': [location],
                               'Cost Center': [cost_center],
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
                                'Special Holiday Excess 8Hrs': [Specialholiday_Restday_Overtime_Excess ],
                                'Legal Holiday':[legalholiday],
                                'Legal Holiday_1st 8hrs':[legalholiday_Overtime],
                                'Legal Holiday_Excess of 8hrs':[legalholiday_Overtime_Excess],
                                'Legal Holiday Falling on Rest Day_1st 8hrs':[legalholiday_RestDay_Overtime],
                                'Legal Holiday Falling on Rest Day_Excess of 8hrs':[legalholiday_RestDay_Overtime_Excess],
                                'Night Differential Regular Days_1st 8hrs':[RegularDay_Night_8hours],
                                'Night Differential Regular Days_Excess of 8hrs':[RegularDay_Night_8hours_Excess],
                                'Night Differential Falling on Rest Day_1st 8hrs':[RegularDay_RestDay_Night_8hours],
                                'Night Differential Falling on Rest Day_Excess of 8hrs': [RegularDay_RestDay_Night_8hours_Excess],
                                'Night Differential Falling on SPHOL rest day 1st 8 hr': [Specialholiday_RestDay_Night_8hours],
                                'Night Differential SH falling on RD_EX8':[Specialholiday_RestDay_Night_8hours_Excess],
                                'Night Differential on Legal Holidays falling on Rest Days':[legalholiday_rest_Night_diffence],
                                'Night Differential on Legal Holidays_1st 8hrs':[legalholiday_Night_8hours],
                                'Night Differential on Legal Holidays_Excess of 8hrs':[legalholiday_Night_8hours_Excess],
                                'Night Differential falling on Special Holiday':[Specialholiday_Night_diffence],
                                'Night Differential SH_EX8':[Specialholiday_Night_8hours_Excess]})

        try:
            existing_df = pd.read_csv('DTR_Summary.csv')
        except FileNotFoundError:

            existing_df = pd.DataFrame(columns=['Employee_Code',
                                                'Employee_Name',
                                                'Position',
                                                'Location',
                                                'Cost Center',
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
                                                'Special Holiday Excess 8Hrs',
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
        dfexcel = pd.read_csv('DTR_Summary.csv')
    except FileNotFoundError:
        dfexcel = pd.DataFrame(columns=['Employee_Code',
                                            'Employee_Name',
                                            'Position',
                                            'Location',
                                            'Cost Center',
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
                                            'Special Holiday Excess 8Hrs',
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


    # Step 2: Load the existing Excel file
    template_file_path = 'excel_temp.xlsx'  # Replace with the path to your template Excel file
    template_wb = load_workbook(template_file_path)
    template_ws = template_wb.active

    # Step 3: Transfer the CSV values to specific cells in the Excel file
    # Specify the target cells where you want to transfer the values
    cell_mapping = {
        'Employee_Code': 'A6',
        'Employee_Name': 'B6',
        'Location': 'D6',
        'Position': 'F6',
        'Cost Center': 'G6',
        'No. of Working Days': 'J6',
        'Number of Working Hours': 'K6',
        'Tardiness/Undertime': 'L6',
        'No. of Days Absent': 'M6',
        'ROT_125': 'N6',
        'Regular OT 100': 'O6',
        'Rest Day': 'P6',
        'RestDay Overtime for the 1st 8hrs': 'Q6',
        'Rest Day Overtime in Excess of 8hrs': 'R6',
        'Special Holiday': 'S6',
        'Special Holiday_1st 8hrs': 'T6',
        'Special Holiday_Excess of 8hrs': 'U6',
        'Special Holiday Falling on restday 1st 8hrs': 'V6',
        'Special Holiday Excess 8Hrs': 'W6',
        'Legal Holiday': 'X6',
        'Legal Holiday_1st 8hrs': 'Y6',
        'Legal Holiday_Excess of 8hrs': 'Z6',
        'Legal Holiday Falling on Rest Day_1st 8hrs': 'AA6',
        'Legal Holiday Falling on Rest Day_Excess of 8hrs': 'AB6',
        'Night Differential Regular Days_1st 8hrs': 'AC6',
        'Night Differential Regular Days_Excess of 8hrs': 'AD6',
        'Night Differential Falling on Rest Day_1st 8hrs': 'AE6',
        'Night Differential Falling on Rest Day_Excess of 8hrs': 'AF6',
        'Night Differential Falling on SPHOL rest day 1st 8 hr': 'AG6',
        'Night Differential on Legal Holidays falling on Rest Days': 'AH6',
        'Night Differential on Legal Holidays_1st 8hrs': 'AI6',
        'Night Differential on Legal Holidays_Excess of 8hrs': 'AJ6',
        'Night Differential SH falling on RD_EX8': 'AK6',
        'Night Differential falling on Special Holiday': 'AL6',
        'Night Differential SH_EX8': 'AM6'
    }

    for column, cell in cell_mapping.items():
        if column in dfexcel.columns:
            values = dfexcel[column].values
            for row, value in enumerate(values, start=6):
                template_ws[cell.replace('6', str(row))] = value

    # Step 4: Save the modified Excel file
    temp_dir = tempfile.gettempdir()
    output_file_path = os.path.join(temp_dir, 'DTR_excel.xls')  # Replace with the desired path for the modified file
    template_wb.save(output_file_path)




    # Download the new CSV file
    return send_file(output_file_path, as_attachment=True)


if __name__ == '__main__':

    operating_system = platform.system()

    # Set the URL based on the operating system
    if operating_system == 'Windows':
        url = f'http://localhost:{port}'
    elif operating_system == 'Darwin':  # macOS
        url = f'http://localhost:{port}'
    elif operating_system == 'Linux':
        url = f'http://localhost:{port}'
    else:
        raise NotImplementedError(f'Unsupported operating system: {operating_system}')

    # Create a new thread to open the browser
    browser_thread = threading.Thread(target=open_browser, args=(url,))
    browser_thread.start()


    app.run(use_reloader=False, port=port)

