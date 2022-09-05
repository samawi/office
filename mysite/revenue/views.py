from django.shortcuts import render
from django.http import HttpResponse

import openpyxl
import pandas as pd
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta
from itertools import islice




# open excel file
excel_file = "tenancy_list_30aug2022.xlsx"
wb = openpyxl.load_workbook(excel_file, data_only=True)

# getting all sheets
sheets = wb.sheetnames

# getting a particular sheet
sheet_params = wb["params"]
sheet_data = wb["tenant"]
sheet_rental = wb["rental"]
sheet_yearly_rental = wb["yearly_rental"]
sheet_sc = wb["sc"]
sheet_yearly_sc = wb["yearly_sc"]
sheet_total = wb["total_rev"]
sheet_occ = wb["occ_rate"]

# reading parameters
date_start = sheet_params["C4"].value
date_end = sheet_params['C5'].value
area_rentable_office = sheet_params["C6"].value
report_sum = sheet_params["C7"].value

# generate reporting daterange based on start and end dates
my_reporting_date_range = pd.date_range(start=date_start, end=date_end, freq='MS')

# Function: Read a sheet into dataframe
def read_worksheet_into_dataframe(sheet_data):
    data = sheet_data.values
    cols = next(data)[1:]
    data = list(data)
    #idx = [r[0] for r in data]
    data = (islice(r, 1, None) for r in data)
    df_data = pd.DataFrame(data, columns=cols)

    # drop rows where column Tenant is Null/NA
    df_data = df_data[df_data['Tenant'].notna()]
    return df_data

# Function: Calculate revenue
def calculate_revenue(date_start, date_end, df_data):
    # Create report dataframe using datetime index for period
    df_report_rental_charge = pd.DataFrame(my_reporting_date_range, columns=['date'])   # Rental revenue
    df_report_sc_charge = pd.DataFrame(my_reporting_date_range, columns=['date'])       # Service charge revenue
    df_report_occupancy = pd.DataFrame(my_reporting_date_range, columns=['date'])       # Occupancy rate

    # Set 'date' as the index
    df_report_rental_charge = df_report_rental_charge.set_index('date')
    df_report_sc_charge = df_report_sc_charge.set_index('date')
    df_report_occupancy = df_report_occupancy.set_index('date')

    # Calculate how much is SC rate based on months and years
    # April 2022: 75000
    # April 2024: 80000
    # April 2026: 85000

    sc_data = [
        [datetime(2018, 4, 1), 65000.00],
        [datetime(2020, 4, 1), 70000.00],
        [datetime(2022, 4, 1), 70000.00],
        [datetime(2024, 4, 1), 75000.00],
        [datetime(2026, 4, 1), 80000.00],
        [datetime(2028, 4, 1), 85000.00]
    ]
    substring = '-'
    for index, row in df_data.iterrows():
        vacant = False
    
        data_area = row['Area']
        data_rental_rate = row['Rental_Rate']
        data_sc_rate = row['SC_Rate']

        if (data_area == None) or (substring in str(data_area)):
            data_area = 0.0

        if (data_rental_rate == None) or (substring in str(data_rental_rate)):
            data_rental_rate = 0.0

        if (data_sc_rate == None) or (substring in str(data_sc_rate)):
            data_sc_rate = 0.0

    # calculate monthly rental and service charge
        calc_rental_charge = data_area * data_rental_rate
        calc_service_charge = data_area * data_sc_rate

        start = row["Start"]
        end = row["End"]

        if (pd.isna(end)) or (end == '-'):
            vacant = True
            end = date_end

        if not vacant:
            if (pd.isna(start)) or (start == '-'):
                start = date_start

            str_level = str(row['Floor']).split('.')[0]
            str_zone = str(row['Zone'])
    
            if (str_level == 'None'):
                str_level = 'NA'
            if (str_zone == 'None'):
                str_zone = 'NA'

            # join them
            str_sep = '-'
            str_temps = [str(str_level), str(str_zone), str(index)]
            str_column_name = str_sep.join(str_temps)

            tenant_date_range = pd.date_range(start=start, end=end, freq='MS')

            # generate rental charge
            df_tenant_rental_charge = pd.DataFrame(tenant_date_range, columns=['date']) # create new df with only 1 column called 'date' and fill with the date range from lcd to led
            df_tenant_rental_charge = df_tenant_rental_charge.set_index('date') # set 'date' as the index
            df_tenant_rental_charge[str_column_name] = calc_rental_charge # add a new column called <num> and fill all with the rental_charge value
            df_report_rental_charge = df_report_rental_charge.join(df_tenant_rental_charge, how="left")

            # generate service charge report
            df_sc_data = pd.DataFrame(data=sc_data, columns=['date', 'sc']).set_index('date')
            sc_date_range = pd.date_range(start=df_sc_data.index.values.min(), end=df_sc_data.index.values.max(), freq='D')
            df_sc = pd.DataFrame(sc_date_range, columns=['date'])
            df_sc = df_sc.set_index('date')

            for mydate in df_sc_data.index:
                date1 = mydate
                date2 = mydate+relativedelta(years=2)-relativedelta(days=1)
            #print('{0} - {1}'.format(date1.strftime('%Y-%m-%d'), date2.strftime('%Y-%m-%d')))
                if data_sc_rate == 0:
                    df_sc.loc[date1.strftime('%Y-%m-%d'):date2.strftime('%Y-%m-%d'), str_column_name] = 0
                elif data_sc_rate == 84100:
                    df_sc.loc[date1.strftime('%Y-%m-%d'):date2.strftime('%Y-%m-%d'), str_column_name] = 84100
                else:
                    df_sc.loc[date1.strftime('%Y-%m-%d'):date2.strftime('%Y-%m-%d'), str_column_name] = df_sc_data.loc[mydate, 'sc']


            df_sc = df_sc * data_area
            tenant_date_range = pd.date_range(start=start, end=end, freq='MS')

            df_tenant_service_charge = pd.DataFrame(tenant_date_range, columns=['date'])
            df_tenant_service_charge = df_tenant_service_charge.set_index('date')
            df_tenant_service_charge = df_tenant_service_charge.join(df_sc, how='left')
            df_report_sc_charge = df_report_sc_charge.join(df_tenant_service_charge, how="left")

            # generate occupancy report
            df_tenant_occupancy = pd.DataFrame(tenant_date_range, columns=['date'])
            df_tenant_occupancy = df_tenant_occupancy.set_index('date')
            df_tenant_occupancy[str_column_name] = data_area
            df_report_occupancy = df_report_occupancy.join(df_tenant_occupancy, how="left")

        # Clean dataframes
        df_report_rental_charge.fillna(0, inplace=True)
        df_report_sc_charge.fillna(0, inplace=True)
        df_report_occupancy.fillna(0, inplace=True)

        # sum each rows
        df_report_rental_charge['sum'] = df_report_rental_charge.sum(axis=1)
        df_report_sc_charge['sum'] = df_report_sc_charge.sum(axis=1)
        df_report_occupancy['sum'] = df_report_occupancy.sum(axis=1)
        # create summary report
        df_sum = pd.DataFrame()
        df_sum['Rental'] = df_report_rental_charge.resample(report_sum).sum()['sum']
        df_sum['SC'] = df_report_sc_charge.resample(report_sum).sum()['sum']
        df_sum['Total'] = df_sum.sum(axis=1)

        df_sum['Occ'] = df_report_occupancy.resample(report_sum).sum()['sum']
        df_sum['OccPct'] = df_sum['Occ']/area_rentable_office
    return df_sum, df_report_rental_charge,df_report_sc_charge,df_report_occupancy

df_data = read_worksheet_into_dataframe(sheet_data)
df_sum, df_report_rental_charge, df_report_sc_charge, df_report_occupancy = calculate_revenue(date_start, date_end , df_data)

# Create your views here.
def index(request):
    my_object = df_sum.to_html(classes='table table-striped')
    return HttpResponse(my_object)


