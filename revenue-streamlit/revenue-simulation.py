import streamlit as st
import pandas as pd
import plotly.express as px
import altair as alt
import numpy as np
from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode

import openpyxl
import pandas as pd
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta
from itertools import islice

# open excel file
excel_file = "tenancy_list_05SEP2022.xlsx"
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
#report_sum = sheet_params["C7"].value

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
def calculate_revenue(date_start, date_end, df_data, product_type, report_sum):
    # Create report dataframe using datetime index for period
    df_report_rental_charge = pd.DataFrame(my_reporting_date_range, columns=['date'])   # Rental revenue
    df_report_sc_charge = pd.DataFrame(my_reporting_date_range, columns=['date'])       # Service charge revenue
    df_report_sc_charge_rate = pd.DataFrame(my_reporting_date_range, columns=['date'])       # Service charge revenue
    df_report_occupancy = pd.DataFrame(my_reporting_date_range, columns=['date'])       # Occupancy rate

    # Set 'date' as the index
    df_report_rental_charge = df_report_rental_charge.set_index('date')
    df_report_sc_charge = df_report_sc_charge.set_index('date')
    df_report_sc_charge_rate = df_report_sc_charge_rate.set_index('date')
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
        [datetime(2027, 4, 1), 80000.00],
        [datetime(2030, 4, 1), 85000.00],
        [datetime(2033, 4, 1), 90000.00],
        [datetime(2027, 4, 1), 95000.00],
    ]

    # select rows according to product_type
    options = [product_type]
    df_by_product = df_data[df_data['Product_Type'].isin(options)]

    substring = '-'
    for index, row in df_by_product.iterrows():
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
            # df_sc_data = pd.DataFrame(data=sc_data, columns=['date', 'sc']).set_index('date')
            df_sc_data = pd.DataFrame(data=sc_data, columns=['date', 'sc'])
            df_sc_data.reset_index(inplace=True)
            sc_date_range = pd.date_range(start=df_sc_data.date.values.min(), end=df_sc_data.date.values.max(), freq='D')
            df_sc = pd.DataFrame(sc_date_range, columns=['date'])
            df_sc = df_sc.set_index('date')

            length = len(df_sc_data.index)
            for i in range(length):
                if i < length-1:
                    date1 = df_sc_data.iloc[i,1]
                    date2 = df_sc_data.iloc[i+1,1]
                    sc_rate = df_sc_data.iloc[i,2]
                    # print(f"{date1} to {date2} is {sc_rate}" )
                    if data_sc_rate == 0:
                        df_sc.loc[date1.strftime('%Y-%m-%d'):date2.strftime('%Y-%m-%d'), str_column_name] = 0
                    elif data_sc_rate == 84100:
                        df_sc.loc[date1.strftime('%Y-%m-%d'):date2.strftime('%Y-%m-%d'), str_column_name] = 84100
                    else:
                        df_sc.loc[date1.strftime('%Y-%m-%d'):date2.strftime('%Y-%m-%d'), str_column_name] = df_sc_data.loc[i, 'sc']

            df_sc_rate_temp = df_sc
            df_sc = df_sc * data_area
            tenant_date_range = pd.date_range(start=start, end=end, freq='MS')

            df_tenant_service_charge = pd.DataFrame(tenant_date_range, columns=['date'])
            df_tenant_service_charge = df_tenant_service_charge.set_index('date')
            df_tenant_service_charge = df_tenant_service_charge.join(df_sc, how='left')
            df_report_sc_charge = df_report_sc_charge.join(df_tenant_service_charge, how="left")

            df_tenant_service_charge_rate = pd.DataFrame(tenant_date_range, columns=['date'])
            df_tenant_service_charge_rate = df_tenant_service_charge_rate.set_index('date')
            df_tenant_service_charge_rate = df_tenant_service_charge_rate.join(df_sc_rate_temp, how='left')
            df_report_sc_charge_rate = df_report_sc_charge_rate.join(df_tenant_service_charge_rate, how="left")

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

    df_sum['Occ'] = df_report_occupancy.resample(report_sum).mean()['sum']
    df_sum['OccPct'] = df_sum['Occ']/area_rentable_office
    df_sum.reset_index(inplace=True)
    return df_sum, df_report_rental_charge, df_report_sc_charge, df_report_occupancy, df_report_sc_charge_rate

df_data = read_worksheet_into_dataframe(sheet_data)
df_sum, df_report_rental_charge, df_report_sc_charge, df_report_occupancy, df_report_sc_charge_rate = calculate_revenue(date_start, date_end , df_data, 'Office', 'Y')

# df_sum.to_excel('test1.xlsx')

st.set_page_config(layout = "wide")

fig = px.bar(
    df_sum,
    x = 'date',
    y = 'OccPct',
    title = "Occupancy Rates",
)
fig.update_layout(xaxis_tickangle = -45)
st.plotly_chart(fig, use_container_width = True)

fig = px.bar(
    df_sum,
    x = 'date',
    y = ["Rental", "SC"],
    title = "Revenue",
    labels = {'Rental':'Rp'},
)
fig.update_layout(xaxis_tickangle = -45)
st.plotly_chart(fig, use_container_width = True)

format_mapping = {
    "Rental": "Rp {:,.2f}",
    "SC": "Rp {:,.2f}",
    "Total": "Rp {:,.2f}",
    "Occ": "{:,.2f} m2",
    "OccPct": "{:,.2%}",
}

# print dataframe
"The Energy | Office Rental"
df_sum_styled = df_sum.style.format(format_mapping)
st.table(df_sum_styled)

# B/W Heatmap of occupancy
df = df_report_occupancy.drop(['sum'], axis = 1)
df.mask(df > 0, 1, inplace =True)

fig4 = px.imshow(df.loc['2023'].iloc[:,:], color_continuous_scale="gray")
fig4.update_traces(xgap = 1, ygap = 1)
st.plotly_chart(fig4, use_container_width = True)

# Heatmap chart
df = df_report_rental_charge.loc['2030']
df = df.drop(['sum'], axis = 1)
df = df.T
fig2 = px.density_heatmap(df)
st.plotly_chart(fig2, use_container_width = True)

fig3 = px.colors.qualitative.swatches()
st.plotly_chart(fig3)
