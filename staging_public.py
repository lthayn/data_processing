import pandas as pd
from functools import reduce
import re

folder = r"XXX"

file = r"\XXX.xlsx"

s_data = pd.read_excel(folder + file, sheet_name='S', converters={'Subheading': str})
s_plus_data = pd.read_excel(folder + file, sheet_name='S+', converters={'Subheading': str})
s_plus_data = s_plus_data[s_plus_data[2020] !='Free']

#print(s_plus_data.columns)

def nine(year):
    #this is to identify and prepare the rows that refer to the hts for duty rate computation purposes
    #gets rows that start with "See"
    mask = s_plus_data[year].str.startswith("See", na=False)
    nine_plus = s_plus_data[mask]

    #just the column for that year
    nine_plus = nine_plus.loc[:, nine_plus.columns.intersection(['Subheading', year])]

    #sets up all the columns for that year, drops unneeded year column at the end
    nine_plus['r' + str(year) + '_rate_type_code'] = "9"
    nine_plus['r' + str(year) + '_text_rate'] = nine_plus[year] + " (S+)"
    nine_plus['r' + str(year) + '_ad_val_rate'] = 9999.99999
    nine_plus['r' + str(year) + '_specific_rate'] = 9999.99999
    nine_plus['r' + str(year) + '_other_rate'] = 9999.99999
    nine_plus = nine_plus.drop([year], axis=1)
    return nine_plus

#list of year iterators
years = [2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]

#this creates a dictionary where each year of data is stored in a separate part of the dictionary
year_9 = {}

for i in years:
    year_9['nine_' + str(i)] = nine(i)

locals().update(year_9)

#outer join on all the rows that refer to the hts for duty rate computation purposes
dataframes = [nine_2020, nine_2021, nine_2022, nine_2023, nine_2024, nine_2025, nine_2026, nine_2027, nine_2028, nine_2029, nine_2030]

nine_combined = reduce(lambda  left,right: pd.merge(left,right,on=['Subheading'], how='outer'), dataframes )


def four_dollar(year):
    #this is to identify and prepare the rows that refer to (Specific rate*Q1) + (Ad Valorem rate*Value)
    #when the specific rate is in dollars and not cents
    p = re.compile('[$].+[/][a-zA-Z]+[ + ].+[%]')

    four = s_plus_data[s_plus_data[year].str.contains(p, na=False)]
    four = four.loc[:, four.columns.intersection(['Subheading', year])]

    # new data frame with split value columns
    new = four[year].str.split("/kg", n=1, expand=True)

    # setting up new columns
    four["r" + str(year) + "_rate_type_code"] = "4"
    four["r" + str(year) + "_text_rate"] = four[year]
    # ad_val_rate
    ad_val = new[1]
    ad_val = ad_val.map(lambda x: x.lstrip('+ ').rstrip('%'))
    ad_val = pd.to_numeric(ad_val)
    four["r" + str(year) + "_ad_val_rate"] = ad_val.div(100)
    # specific rate
    specific = new[0]
    specific = specific.map(lambda x: x.lstrip('$'))
    four["r" + str(year) + "_specific_rate"] = pd.to_numeric(specific)
    # other
    four["r" + str(year) + "_other_rate"] = 0

    four = four.drop([year], axis=1)

    return four

#list of year iterators
years = [2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]

#this creates a dictionary where each year of data is stored in a separate part of the dictionary
year_4 = {}

for i in years:
    year_4['four_d_' + str(i)] = four_dollar(i)

locals().update(year_4)

#outer join on all the rows that refer to the hts for duty rate computation purposes
dataframes = [four_d_2020, four_d_2021, four_d_2022, four_d_2023, four_d_2024, four_d_2025, four_d_2026, four_d_2027, four_d_2028, four_d_2029, four_d_2030]

four_d_combined = reduce(lambda  left,right: pd.merge(left,right,on=['Subheading'], how='outer'), dataframes)

def four_cent(year):
    #this is to identify and prepare the rows that refer to (Specific rate*Q1) + (Ad Valorem rate*Value)
    #when the specific rate is listed in cents and not dollars
    p = re.compile('.+[ cents]{1}[/][a-zA-Z]+[ + ].+[%]')

    four = s_plus_data[s_plus_data[year].str.contains(p, na=False)]
    four = four.loc[:, four.columns.intersection(['Subheading', year])]

    # new data frame with split value columns
    new = four[year].str.split("/", n=1, expand=True)

    # setting up new columns
    four["r" + str(year) + "_rate_type_code"] = "4"
    four["r" + str(year) + "_text_rate"] = four[year]
    # ad_val_rate
    ad_val = new[1]
    ad_val = ad_val.map(lambda x: x.lstrip('liter + ').lstrip('kg + ').lstrip('of total sugars + ').rstrip('%'))
    ad_val = pd.to_numeric(ad_val)
    four["r" + str(year) + "_ad_val_rate"] = ad_val.div(100)
    # specific rate
    specific = new[0]
    specific = specific.map(lambda x: x.rstrip('cents'))
    specific = pd.to_numeric(specific)
    four["r" + str(year) + "_specific_rate"] = specific.div(100)
    # other
    four["r" + str(year) + "_other_rate"] = 0

    four = four.drop([year], axis=1)

    return four

#this creates a dictionary where each year of data is stored in a separate part of the dictionary
year_4_c = {}

for i in years:
    year_4_c['four_c_' + str(i)] = four_cent(i)

locals().update(year_4_c)

#outer join on all the rows that refer to the hts for duty rate computation purposes
dataframes = [four_c_2020, four_c_2021, four_c_2022, four_c_2023, four_c_2024, four_c_2025, four_c_2026, four_c_2027, four_c_2028, four_c_2029, four_c_2030]

four_c_combined = reduce(lambda  left,right: pd.merge(left,right,on=['Subheading'], how='outer'), dataframes)

#this appends the separate processed data together
combined = nine_combined.append(four_d_combined, ignore_index=True)
combined = combined.append(four_c_combined, ignore_index=True)

#create a new dataframe i'm drawing from where the processed rows no longer appear.
#this keeps those hts8 numbers in the original s_plus_data
smaller_s_plus = s_plus_data.merge(combined, on=['Subheading'], how='outer', indicator=True).loc[
    lambda x: x['_merge'] == 'left_only']
smaller_s_plus = smaller_s_plus.drop(['_merge'], axis=1)

def one_cent(year, data_set):
    #this is to identify and prepare the rows that refer to Specific rate*Q1 when specific rate is listed in cents
    p = re.compile('.+[ cents]{1}[/][a-zA-Z]+')

    one = data_set[data_set[year].str.contains(p, na=False)]
    one = one.loc[:, one.columns.intersection(['Subheading', year])]

    # setting up new columns
    one["r" + str(year) + "_rate_type_code"] = "1"
    one["r" + str(year) + "_text_rate"] = one[year]
    # ad_val_rate
    one["r" + str(year) + "_ad_val_rate"] = 0
    # specific rate
    specific = one[year]
    specific = specific.map(lambda x: x.rstrip('cents/liter').rstrip('cents/kg'))
    specific = pd.to_numeric(specific)
    one["r" + str(year) + "_specific_rate"] = specific.div(100)
    # other
    one["r" + str(year) + "_other_rate"] = 0

    one = one.drop([year], axis=1)

    return one

#this creates a dictionary where each year of data is stored in a separate part of the dictionary
year_1_c = {}

for i in years:
    year_1_c['one_c_' + str(i)] = one_cent(i, smaller_s_plus)

locals().update(year_1_c)

#outer join on all the rows that refer to the hts for duty rate computation purposes
dataframes = [one_c_2020, one_c_2021, one_c_2022, one_c_2023, one_c_2024, one_c_2025, one_c_2026, one_c_2027, one_c_2028, one_c_2029, one_c_2030]

one_c_combined = reduce(lambda  left,right: pd.merge(left,right,on=['Subheading'], how='outer'), dataframes)


#this appends the separate processed data together
combined = combined.append(one_c_combined, ignore_index=True)

#create a new dataframe i'm drawing from where the processed rows no longer appear.
#this keeps those hts8 numbers in the original s_plus_data
smaller_s_plus = s_plus_data.merge(combined, on=['Subheading'], how='outer', indicator=True).loc[
    lambda x: x['_merge'] == 'left_only']
smaller_s_plus = smaller_s_plus.drop(['_merge'], axis=1)


def one_dollar(year, data_set):
    #this is to identify and prepare the rows that refer to Specific rate*Q1 when specific rate is listed in dollars
    p = re.compile('[$]{1}.+[/]{1}[a-zA-Z]+')

    one = data_set[data_set[year].str.contains(p, na=False)]
    one = one.loc[:, one.columns.intersection(['Subheading', year])]

    # setting up new columns
    one["r" + str(year) + "_rate_type_code"] = "1"
    one["r" + str(year) + "_text_rate"] = one[year]
    # ad_val_rate
    one["r" + str(year) + "_ad_val_rate"] = 0
    # specific rate
    specific = one[year]
    specific = specific.map(lambda x: x.lstrip('$').rstrip('/kg'))
    one["r" + str(year) + "_specific_rate"] = pd.to_numeric(specific)
    # other
    one["r" + str(year) + "_other_rate"] = 0

    one = one.drop([year], axis=1)

    return one

#this creates a dictionary where each year of data is stored in a separate part of the dictionary
year_1_d = {}

for i in years:
    year_1_d['one_d_' + str(i)] = one_dollar(i, smaller_s_plus)

locals().update(year_1_d)

#outer join on all the rows that refer to the hts for duty rate computation purposes
dataframes = [one_d_2020, one_d_2021, one_d_2022, one_d_2023, one_d_2024, one_d_2025, one_d_2026, one_d_2027, one_d_2028, one_d_2029, one_d_2030]

one_d_combined = reduce(lambda  left,right: pd.merge(left,right,on=['Subheading'], how='outer'), dataframes)


#this appends the separate processed data together
combined = combined.append(one_d_combined, ignore_index=True)

#create a new dataframe i'm drawing from where the processed rows no longer appear.
#this keeps those hts8 numbers in the original s_plus_data
smaller_s_plus = s_plus_data.merge(combined, on=['Subheading'], how='outer', indicator=True).loc[
    lambda x: x['_merge'] == 'left_only']
smaller_s_plus = smaller_s_plus.drop(['_merge'], axis=1)


def seven(year, data_set):
    #this is to identify and prepare the rows that refer to Ad Valorem rate*Value
    seven = data_set
    #drops all columns but year and subheading
    seven = seven.loc[:, seven.columns.intersection(['Subheading', year])]
    # setting up new columns
    seven["r" + str(year) + "_rate_type_code"] = "7"
    text_r = seven[year]
    text_r = text_r.mul(100)
    seven["r" + str(year) + "_text_rate"] = text_r.astype(str) + "%"
    # ad_val_rate
    ad_val = seven[year]
    seven["r" + str(year) + "_ad_val_rate"] = pd.to_numeric(ad_val)
    # specific rate
    seven["r" + str(year) + "_specific_rate"] = 0
    # other
    seven["r" + str(year) + "_other_rate"] = 0

    seven = seven.drop([year], axis=1)

    return seven

#this creates a dictionary where each year of data is stored in a separate part of the dictionary
year_7 = {}

for i in years:
    year_7['seven_' + str(i)] = seven(i, smaller_s_plus)

locals().update(year_7)

#outer join on all the rows that refer to the hts for duty rate computation purposes
dataframes = [seven_2020, seven_2021, seven_2022, seven_2023, seven_2024, seven_2025, seven_2026, seven_2027, seven_2028, seven_2029, seven_2030]

seven_combined = reduce(lambda  left,right: pd.merge(left,right,on=['Subheading'], how='outer'), dataframes)

#this appends the separate processed data together
combined = combined.append(seven_combined, ignore_index=True)


#output into excel
combined.to_excel(r"XXX\XXX.xlsx")
