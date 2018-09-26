import calendar
import xlsxwriter
import pandas as pd
from Config import config


def generate_report_xlsx(dates_list, frame_dates_list, output_name):

    workbook = xlsxwriter.Workbook(f'output/{output_name}.xlsx')
    worksheet = workbook.add_worksheet()
    text_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter'})

    row = 1
    col = 1
    for irow in dates_list:
        worksheet.write_rich_string(row, col, irow, text_format)
        row = row + 1

    row = 1
    col += 1
    for irow in frame_dates_list:
        worksheet.write_rich_string(row, col, irow, text_format)
        row = row + 1

    workbook.close()


def get_starts_days(all_dates, day_name):

    all_wednesday = []
    for date in all_dates:
        if date.weekday_name == day_name:
            all_wednesday.append(date)

    return all_wednesday


def get_wetaca_dates(wednesday_list):

    dates_dict = {}
    for index, date in enumerate(wednesday_list):

        dates_dict['Week_{}'.format(index)] = {

            "pay_date": date,
            "delyvery_date": date + pd.DateOffset(4),
            "start_date": date + pd.DateOffset(5),
            "end_date": date + pd.DateOffset(11)
        }

    return dates_dict


def get_supermarket_dates(saturday_list):

    dates_dict = {}
    for index, date in enumerate(saturday_list):

        dates_dict['Week_{}'.format(index)] = {

            "pay_date": date,
            "start_date": date + pd.DateOffset(1),
            "end_date": date + pd.DateOffset(7)
        }

    return dates_dict


def get_free_time_dates(fridays_list):

    dates_dict = {}
    for index, date in enumerate(fridays_list):

        dates_dict['Week_{}'.format(index)] = {

            "pay_date": date,
            "start_date": date - pd.DateOffset(4),
            "end_date": date + pd.DateOffset(2)
        }

    return dates_dict


def get_family_dates(fridays_list):

    dates_dict = {}
    for index, date in enumerate(fridays_list):

        dates_dict['Week_{}'.format(index)] = {

            "pay_date": date,
            "start_date": date + pd.DateOffset(1),
            "end_date": date + pd.DateOffset(7)
        }

    return dates_dict


def wetaca_dates():

    start_date = pd.to_datetime(config.START_DATE)
    end_date = pd.to_datetime(config.END_DATE)

    all_dates = pd.date_range(start=start_date, end=end_date)

    wednesday_list = get_starts_days(all_dates, "Wednesday")

    dates_dict = get_wetaca_dates(wednesday_list)

    df = pd.DataFrame.from_dict(dates_dict, orient='index')

    dates_list = []
    frame_dates_list = []
    for row in df.iterrows():
        pay_weekday_name = row[1]["pay_date"].weekday_name
        pay_day_number = row[1]['pay_date'].day
        pay_month_name = calendar.month_name[row[1]['pay_date'].month]
        delyvery_weekday_name = row[1]['delyvery_date'].weekday_name
        delyvery_day_number = row[1]['delyvery_date'].day
        delyvery_month_name = calendar.month_name[row[1]['delyvery_date'].month]

        start_day_week_day_name = row[1]['start_date'].weekday_name
        start_day_number = row[1]['start_date'].day
        start_frame_month_name = calendar.month_name[row[1]['start_date'].month]
        end_day_week_day_name = row[1]['end_date'].weekday_name
        end_day_number = row[1]['end_date'].day
        end_frame_month_name = calendar.month_name[row[1]['end_date'].month]

        dates_list.append(
            f"Pay day:\n{pay_weekday_name} {pay_day_number} {pay_month_name}\n"
            f"Delyvery date:\n{delyvery_weekday_name} {delyvery_day_number} {delyvery_month_name}"
        )

        frame_dates_list.append(
            f"{start_day_week_day_name} {start_day_number} {start_frame_month_name} - "
            f"{end_day_week_day_name} {end_day_number} {end_frame_month_name}"
        )

    return dates_list, frame_dates_list


def supermarket_dates():

    start_date = pd.to_datetime(config.START_DATE)
    end_date = pd.to_datetime(config.END_DATE)

    all_dates = pd.date_range(start=start_date, end=end_date)

    saturdays_list = get_starts_days(all_dates, 'Saturday')

    dates_dict = get_supermarket_dates(saturdays_list)

    df = pd.DataFrame.from_dict(dates_dict, orient='index')

    dates_list = []
    frame_dates_list = []
    for row in df.iterrows():
        pay_weekday_name = row[1]["pay_date"].weekday_name
        pay_day_number = row[1]['pay_date'].day
        pay_month_name = calendar.month_name[row[1]['pay_date'].month]

        start_day_week_day_name = row[1]['start_date'].weekday_name
        start_day_number = row[1]['start_date'].day
        start_frame_month_name = calendar.month_name[row[1]['start_date'].month]
        end_day_week_day_name = row[1]['end_date'].weekday_name
        end_day_number = row[1]['end_date'].day
        end_frame_month_name = calendar.month_name[row[1]['end_date'].month]

        dates_list.append(
            f"Pay day:\n{pay_weekday_name} {pay_day_number} {pay_month_name}\n"
        )

        frame_dates_list.append(
            f"{start_day_week_day_name} {start_day_number} {start_frame_month_name} - "
            f"{end_day_week_day_name} {end_day_number} {end_frame_month_name}"
        )

    return dates_list, frame_dates_list


def free_time_dates():

    start_date = pd.to_datetime(config.START_DATE)
    end_date = pd.to_datetime(config.END_DATE)

    all_dates = pd.date_range(start=start_date, end=end_date)

    fridays_list = get_starts_days(all_dates, "Friday")

    dates_dict = get_free_time_dates(fridays_list)

    df = pd.DataFrame.from_dict(dates_dict, orient='index')

    dates_list = []
    frame_dates_list = []
    for row in df.iterrows():
        pay_weekday_name = row[1]["pay_date"].weekday_name
        pay_day_number = row[1]['pay_date'].day
        pay_month_name = calendar.month_name[row[1]['pay_date'].month]

        start_day_week_day_name = row[1]['start_date'].weekday_name
        start_day_number = row[1]['start_date'].day
        start_frame_month_name = calendar.month_name[row[1]['start_date'].month]
        end_day_week_day_name = row[1]['end_date'].weekday_name
        end_day_number = row[1]['end_date'].day
        end_frame_month_name = calendar.month_name[row[1]['end_date'].month]

        dates_list.append(
            f"Pay day:\n{pay_weekday_name} {pay_day_number} {pay_month_name}\n"
        )

        frame_dates_list.append(
            f"{start_day_week_day_name} {start_day_number} {start_frame_month_name} - "
            f"{end_day_week_day_name} {end_day_number} {end_frame_month_name}"
        )

    return dates_list, frame_dates_list


def family_dates(report_name):

    if report_name == "supermartket":
        day_name = "Friday"
    elif report_name == "gasoline":
        day_name = "Saturday"
    else:
        print(f'"{report_name}" not supported. execution will stop')
        return 1

    start_date = pd.to_datetime(config.START_DATE)
    end_date = pd.to_datetime(config.END_DATE)

    all_dates = pd.date_range(start=start_date, end=end_date)

    fridays_list = get_starts_days(all_dates, day_name)

    dates_dict = get_family_dates(fridays_list)

    df = pd.DataFrame.from_dict(dates_dict, orient='index')

    dates_list = []
    frame_dates_list = []
    for row in df.iterrows():
        pay_weekday_name = row[1]["pay_date"].weekday_name
        pay_day_number = row[1]['pay_date'].day
        pay_month_name = calendar.month_name[row[1]['pay_date'].month]

        start_day_week_day_name = row[1]['start_date'].weekday_name
        start_day_number = row[1]['start_date'].day
        start_frame_month_name = calendar.month_name[row[1]['start_date'].month]
        end_day_week_day_name = row[1]['end_date'].weekday_name
        end_day_number = row[1]['end_date'].day
        end_frame_month_name = calendar.month_name[row[1]['end_date'].month]

        dates_list.append(
            f"{config.translate_weekday[pay_weekday_name]} {pay_day_number} {config.translate_months[pay_month_name]}\n"
        )

        frame_dates_list.append(
            f"{config.translate_weekday[start_day_week_day_name]} {start_day_number} "
            f"{config.translate_months[start_frame_month_name]} - "
            f"{config.translate_weekday[end_day_week_day_name]} {end_day_number} "
            f"{config.translate_months[end_frame_month_name]}"
        )

    return dates_list, frame_dates_list


if __name__ == '__main__':

    wetaca_dates_list, wetaca_frame_dates_list = wetaca_dates()
    supermarket_dates_list, supermarket_frame_dates_list = supermarket_dates()
    free_time_dates_list, free_time_frame_dates_list = free_time_dates()
    supermarket_family_dates_list, supermarket_family_frame_dates_list = family_dates("supermartket")
    gasoline_family_dates_list, gasoline_family_frame_dates_list = family_dates("gasoline")

    generate_report_xlsx(wetaca_dates_list, wetaca_frame_dates_list, "wetaca_dates")
    generate_report_xlsx(supermarket_dates_list, supermarket_frame_dates_list, "supermarket_dates")
    generate_report_xlsx(free_time_dates_list, free_time_frame_dates_list, "free_time_dates")
    generate_report_xlsx(supermarket_family_dates_list, supermarket_family_frame_dates_list, "supermarket_family")
    generate_report_xlsx(gasoline_family_dates_list, gasoline_family_frame_dates_list, "gasoline_family")
