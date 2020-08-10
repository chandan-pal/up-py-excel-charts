import pandas as pd
import xlsxwriter
import calendar
from datetime import datetime, timedelta


df = pd.read_csv("Log_1_704R.CSV", sep=";", header=None, dtype={
                 6: 'str'}, parse_dates=[10], low_memory=False)
# print(df.head())
# df[10] = df[10].astype('datetime64[ns]')

df_month_basis = [g for n, g in df.groupby(pd.Grouper(key=10, freq="M"))]
for jj in range(len(df_month_basis)):

    df_month = df_month_basis[jj]
    month_year = df_month[10][df_month[10].first_valid_index()].strftime(
        '%B_%Y')

    print('Generating Report for : ' + month_year)

    file_name = 'OutputReport_' + str(month_year) + '.xlsx'
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    df_day_basis = [g for n, g in df_month.groupby(
        pd.Grouper(key=10, freq="D"))]

    for ii in range(len(df_day_basis)):

        df_day = df_day_basis[ii]
        day = df_day[10][df_day[10].first_valid_index()].strftime('%d')
        df_day['Date'] = df_day[10].dt.date
        df_day['Time'] = df_day[10].dt.strftime('%H:%M')
        total_rows, total_cols = df_day.shape
        # print(df_day.head())
        df_day.to_excel(writer, sheet_name=day, startrow=0, index=False, columns=['Date', 'Time', 1, 2, 3, 4, 5, 6, 7, 8, 9])

        # get the control of workbook
        workbook = writer.book

        # data sheet formatting
        table_sheet = writer.sheets[day]
        table_sheet.set_zoom(90)
        table_sheet.add_table(0, 0, total_rows, total_cols-3, {'header_row': False, 'autofilter': 0})
        table_sheet.set_column(0, 0, 12)
        table_sheet.set_column(1, total_cols-3, 8)

        # create a sheet for charts
        chart_sheet = workbook.add_worksheet(day + " Charts")

        # Group data on hour basis
        # df_hour_basis = [i for j, i in df_day.groupby(pd.Grouper(key=10, freq="H"))]
        # for df_hour in df_hour_basis:
        # print(df_hour)
        # pass
        # break

    writer.save()


# Create a new chart object. In this case an embedded chart.
# chart1 = workbook.add_chart({'type': 'line'})

# Configure the first series.
# chart1.add_series({
    # 'name':       '=01!$B$1',
    # 'categories': '=01!$A$2:$A$7',
    # 'values':     '=01!$B$2:$B$7',
# })

# Add a chart title and some axis labels.
# chart1.set_title ({'name': 'Results of sample analysis'})
# chart1.set_x_axis({'name': 'Test number'})
# chart1.set_y_axis({'name': 'Sample length (mm)'})

# Set an Excel chart style. Colors with white outline and shadow.
# chart1.set_style(10)

# Insert the chart into the worksheet (with an offset).
# worksheet.insert_chart('D2', chart1, {'x_offset': 25, 'y_offset': 10})

# workbook.close()
