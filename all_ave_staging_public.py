import pandas as pd
import numpy as np

#reads the file and creates pandas dataframe
excel_file=R"xxxxx.xlsx"
data = pd.read_excel(excel_file, converters={'HTS':str}, index_col=None)

#creates dataframe without rows that have 'Free' in the EIF column
data_2=data[data['EIF'] !='Free']

#drop two unneeded columns
data_2 = data_2.drop(['mfn_text_rate'], axis=1)

#rename EIF as 2020
data_2 = data_2.rename (columns = {'EIF': '2020'})

###this is to create final source column that shows the first year ave became free
#create list of year columns
years=['2021', '2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029']

#create list of column locations for years
year_loc = []

for year in years:
    column = data_2.columns.get_loc(year)
    year_loc.append(column)

#this is loops through the 9 year columns and the first time free appears, that column/year is given as the value in final
final = []
for i in range(len(data_2)):
    if data_2.iloc[i, year_loc[0]] == 'Free':
        final.append(years[0])
    elif data_2.iloc[i, year_loc[1]] == 'Free':
        final.append(years[1])
    elif data_2.iloc[i, year_loc[2]] == 'Free':
        final.append(years[2])
    elif data_2.iloc[i, year_loc[3]] == 'Free':
        final.append(years[3])
    elif data_2.iloc[i, year_loc[4]] == 'Free':
        final.append(years[4])
    elif data_2.iloc[i, year_loc[5]] == 'Free':
        final.append(years[5])
    elif data_2.iloc[i, year_loc[6]] == 'Free':
        final.append(years[6])
    elif data_2.iloc[i, year_loc[7]] == 'Free':
        final.append(years[7])
    elif data_2.iloc[i, year_loc[8]] == 'Free':
        final.append(years[8])
    else:
        final.append('NA')

#assigns final list as final_year column in dataframe
data_2['final_year']=final

def type(data, year):
    # this creates the rate type code, if free it puts 0, else 7
    return np.where(data[year]=='Free', '0', '7')


def ave(data, year):
    #this creates a column with the numerical, fraction version of the text ave
    ave = []
    for row in data[year]:
        if row == 'Free':
            ave.append(0)
        else:
            ave.append(round(float(row.rstrip('%')) / 100.0, 4))

    return ave

#iterate over year columns to create two new columns, rate type and ad vale rate, for year year
years=['2020', '2021', '2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029']

for year in years:
    data_2['r'+year+'_rate_type_code']=type(data_2, year)
    data_2['r'+year+'_ad_val_rate']=ave(data_2, year)

#rename year columsn to rYYYY_text
for year in years:
    data_2 = data_2.rename (columns = {year: 'r'+year+'_text'})

#reorder columns to group years together
#list of column headers in order desired
cols = ['HTS', 'r2020_rate_type_code', 'r2020_text', 'r2020_ad_val_rate', 'r2021_rate_type_code', 'r2021_text',
        'r2021_ad_val_rate', 'r2022_rate_type_code', 'r2022_text', 'r2022_ad_val_rate',
        'r2023_rate_type_code', 'r2023_text', 'r2023_ad_val_rate', 'r2024_rate_type_code', 'r2024_text', 'r2024_ad_val_rate',
        'r2025_rate_type_code', 'r2025_text', 'r2025_ad_val_rate', 'r2026_rate_type_code', 'r2026_text', 'r2026_ad_val_rate',
        'r2027_rate_type_code', 'r2027_text', 'r2027_ad_val_rate', 'r2028_rate_type_code', 'r2028_text', 'r2028_ad_val_rate',
        'r2029_rate_type_code', 'r2029_text', 'r2029_ad_val_rate', 'final_year']

data_3 = data_2[cols]

#output into excel
data_3.to_excel("xxxx.xlsx")