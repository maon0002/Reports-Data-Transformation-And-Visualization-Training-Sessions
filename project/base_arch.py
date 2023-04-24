import numpy as np
import pandas as pd
# import matplotlib.pyplot as plt
# import xlsxwriter
# import datetime as dt

# from xlsxwriter import Workbook


def to_lower(x): return x.lower() if isinstance(x, str) else x


def trim_all_columns(df):
    """
    Trim whitespace from ends of each value across all series in dataframe
    """

    def trim_strings(x): return x.strip() if isinstance(x, str) else x

    return df.applymap(trim_strings)


def unwanted_chars(df):
    """
    Get rid of single quotes and apostrophes
    """
    df = df.replace("'|  |`", "", regex=True)
    return df


def transliteration(df, ser, new_ser):
    """
    Transliterate bg chars into latin when needed
    """
    # create dictionary for transliteration from bg to latin letters
    bg_en = {
        'А': 'A',
        'Б': 'B',
        'В': 'V',
        'Г': 'G',
        'Д': 'D',
        'Е': 'E',
        'Ж': 'Zh',
        'З': 'Z',
        'И': 'I',
        'Й': 'Y',
        'К': 'K',
        'Л': 'L',
        'М': 'M',
        'Н': 'N',
        'О': 'O',
        'П': 'P',
        'Р': 'R',
        'С': 'S',
        'Т': 'T',
        'У': 'U',
        'Ф': 'F',
        'Х': 'H',
        'Ц': 'Ts',
        'Ч': 'Ch',
        'Ш': 'Sh',
        'Щ': 'Sht',
        'Ъ': 'A',
        'Ь': 'Y',
        'Ю': 'Yu',
        'Я': 'Ya',
        'а': 'a',
        'б': 'b',
        'в': 'v',
        'г': 'g',
        'д': 'd',
        'е': 'e',
        'ж': 'zh',
        'з': 'z',
        'и': 'i',
        'й': 'y',
        'к': 'k',
        'л': 'l',
        'м': 'm',
        'н': 'n',
        'о': 'o',
        'п': 'p',
        'р': 'r',
        'с': 's',
        'т': 't',
        'у': 'u',
        'ф': 'f',
        'х': 'h',
        'ц': 'ts',
        'ч': 'ch',
        'ш': 'sh',
        'щ': 'sht',
        'ъ': 'a',
        'ь': 'y',
        'ю': 'yu',
        'я': 'ya',
        ' ': ' ',
        '-': '-'}

    # transliteration char by char
    transliterated_string = ""
    for i in range(len(df[ser])):
        string_value = df.loc[i, ser].strip()
        if string_value and not string_value.isascii():
            for char in string_value:
                if char in bg_en.keys():
                    string_value = string_value.replace(
                        char, bg_en[char])
                    transliterated_string = string_value
                else:
                    transliterated_string = string_value

            df.loc[i, new_ser] = transliterated_string.upper()
            transliterated_string = ""
        else:
            df.loc[i, new_ser] = df.loc[i, ser].upper()
    return df[new_ser]


def nickname(df):
    '''
    By idea the nickname must contain:
    2 letters from the first name,
    3 letters from the last name and
    3 letters from the email (pvt or work w/o only alpha chars)
    '''
    # adding column with both Fname and Lname
    df['employee_names'] = df['first_name'] + " " + df['last_name']

    # adding nickname column
    df['nickname'] = df['first_name'].str[0:2].str.upper() + \
        df['last_name'].str[1:4].str.upper()

    # adding flag and 'X's at the end if the len < 5
    df.loc[df.nickname.str.len() < 5, 'flags'] += '1,'
    df.loc[df.nickname.str.len() < 5, 'nickname'] = df['nickname'].apply(
        lambda x: x + 'X' * (5 - len(x)))

    # check for email columns and take last part of the nickname
    # add flags for the cases
    work_email_validation_filter = (
        df['work_email'].str.contains('@', na=False))
    pvt_email_validation_filter = (df['pvt_email'].str.contains('@', na=False))

    df.loc[~work_email_validation_filter, 'flags'] += "2,"
    df.loc[~pvt_email_validation_filter, 'flags'] += "3,"

    df.loc[~pvt_email_validation_filter &
           ~work_email_validation_filter, 'flags'] += "4,"

    # takes value subtraction from the work or pvt email:
    email_validation_filter = (df['work_email'].str.contains('@', na=False))
    df.loc[email_validation_filter, 'nickname'] += df['work_email'].str.replace(
        "[^A-Za-z]+", "", regex=True).str[1:4].str.upper()

    df.loc[(df['work_email'].str.len() < 1) & (~df['pvt_email'].str.len() < 1) & (
        df['pvt_email'].str.contains('@')), 'nickname'] += df['pvt_email'].str.replace(
        "[^A-Za-z]+", "", regex=True).str[1:4].str.upper()

    df.loc[(df['work_email'].str.len() < 1) & (df['pvt_email'].str.len() < 1),
           'nickname'] = df['nickname'].apply(lambda x: x + 'X' * (8 - len(x)))

    df.loc[df['nickname'].str.len() < 8, 'nickname'] = df['nickname'].apply(
        lambda x: x + 'X' * (8 - len(x)))

    # adding column with both Fname and Lname
    df['employee_names'] = df['employee_names'].str.title()

    return df


def company_substraction(df):
    '''
    1.Check if the meeting was held online or in-person and mark in a separate column
    2.Iterate through the 'type' column and extracts only the company name
    '''

    df.loc[df['type'].str.contains('person|живо', regex=True),
           'short_type'] = "На живо"
    df.loc[df['type'].str.contains('nline|нлайн', regex=True),
           'short_type'] = "Онлайн"
    df.loc[~df['type'].str.contains('nline|нлайн|person|живо',
                                    regex=True), 'flags'] += "7,"

    df.loc[~df['type'].str.len() < 1, 'company'] = df['type'].str.split(
        "[:|/]").str[0].str.upper().str.strip()
    return df


def consultancy_per_emp(df):
    '''
    Add three additional columns for:
        concatenation of employee|company,
        training calculation total by the concat values,
        either the employee have one or more than one consultation
    '''
    # concatenate nickname+company for later count and check if they are equal nicknames from different companies
    df.loc[~df['nickname'].str.len() < 1, 'concat_emp_company'] = df['nickname'].str[:] + \
        "|" + df['company'].str[:]

    # count the frequency of occurrence for every  value (employee)
    df['total_per_emp'] = df.groupby('concat_emp_company')[
        'concat_emp_company'].transform('count')

    # comment if the employee have one or more than one consultancies
    df.loc[df['total_per_emp'].astype(
        int) == 1, 'returns_or_not'] = 'only one session'

    df.loc[df['total_per_emp'].astype(
        int) > 1, 'returns_or_not'] = 'more then one session'
    return df


def phone_validation(df):
    df.loc[df['phone'].str.len() < 1, 'flags'] += "5,"
    df.loc[df['phone'].str.contains(
        '[^\d\+]', regex=True, na=False), 'flags'] += "6,"
    df.loc[df['phone'].str.len().between(1, 8), 'flags'] += "6,"
    return df


def trainers(df):
    '''
    Substract only the non cyrillic names and put them in a separate column
    '''
    df.loc[~df['calendar'].isnull(), 'trainer'] = df['calendar'].str.split(
        '|').str[0].str.strip().str.title()
    return df


def active_contracts(df):
    '''
    Count the consultancies of the company employees based on the consultancy
    date if there's an active contract'
    '''
    # CONCAT the lines for which the training date is between the contract's start and end date
    df.loc[(df['start_time'] >= df['starts']) & (df['start_time'] <=
                                                 df['ends']), 'concat_count'] = df['company'] + "|" + df['nickname']

    # count the company/employees total consultancies for the company active period
    df.loc[(df['start_time'] >= df['starts']) & (df['start_time'] <=
                                                 df['ends']), 'active_training_per_client'] = df.groupby('concat_count')[
        'concat_count'].transform('count')
    # calculate the number of consultancies that can be used
    df.loc[(df['start_time'] >= df['starts']) & (df['start_time'] <=
                                                 df['ends']) & (df['c_per_emp'].between(1, 9998)), 'training_left'] = df[
        'c_per_emp'] - \
        df[
        'active_training_per_client']

    # add flag for employees which have less than 1 training left
    df.loc[(~df['training_left'].isna()) & (df['training_left'] < 2), 'flags'] += '9,'
    return df


def datetime_normalize(df):
    df = df.apply(lambda x: x.strftime("%d-%b-%Y %H:%M:%S"))
    return df


def date_normalize(df):
    df = pd.to_datetime(df).dt.date
    df = df.apply(lambda x: x.strftime("%d-%b-%Y"))
    return df


def dates_diff(ser1, ser2):
    diff = ser2 - ser1
    return diff


def total_trainings_func(df):
    df.insert(10, "language", 'Български', allow_duplicates=False)
    df.insert(10, "status", 'Проведен', allow_duplicates=False)
    return df


def data_transformation(df, df_limit):
    '''
    Transform the monthly and the annual df columns
    '''
    df = df.rename(columns={
        'Start Time': 'start_time', 'End Time': 'end_time', 'Date Scheduled': 'scheduled_on',
        'First Name': 'f_name',
        'Last Name': 'l_name',
        'Phone': 'phone',
        'Email': 'pvt_email',
        'Type': 'type',
        'Calendar': 'calendar',
        'Appointment Price': 'price',
        'Paid?': 'is_paid',
        'Amount Paid Online': 'paid_online',
        'Certificate Code': 'certificate_code',
        'Notes': 'notes',
        'Label': 'label',
        'Scheduled By': 'scheduled_by',
        'Name of the company you work for | Име на фирмата, в която работите': 'company_name',
        'Work email | Служебен имейл': 'work_email',
        'Preferred platforms | Предпочитани платформи': 'preferred_platforms',
        'Appointment ID': 'appointment_id'
    })

    # add column for issues
    df['flags'] = ""
    df['emp_names_input'] = '|' + \
                            df['f_name'] + '|' + df['l_name'] + '|'

    # rough cleaning of the columns data
    df = unwanted_chars(df)
    df = trim_all_columns(df)

    # transliteration of cyrillic to latin chars
    df["first_name"] = transliteration(df, "f_name", "first_name")
    df["last_name"] = transliteration(df, "l_name", "last_name")

    # using first, last name and email parts
    df = nickname(df)

    # get only the company name and if the training was IN PERSON/LIVE or ONLINE
    df = company_substraction(df)

    # merge/vlookup the columns from limitations to the monthly/annual df
    df = pd.merge(
        left=df,
        right=df_limit,
        left_on='company',
        right_on='company',
        how='left')

    # add columns for counting unique emp|company values w/ totals
    df = consultancy_per_emp(df)

    # check phone values
    df = phone_validation(df)

    # substring trainers from calendar via regex.
    df = trainers(df)

    # check if the training date is between the dates when company contract starts and ends
    df = active_contracts(df)

    # create a Month name column
    df["month"] = [pd.Timestamp(x).month_name()
                   for x in df["start_time"]]

    # create a Year name column
    df["year"] = [pd.Timestamp(x).year
                  for x in df["start_time"]]

    # add column with the name of the day when training was take place
    df['dayname'] = df['start_time'].dt.day_name()

    # reformat start_time
    df['training_datetime'] = datetime_normalize(
        df['start_time'])

    # change from datetime to date only
    df['scheduled_date'] = date_normalize(df['scheduled_on'])
    df['training_end'] = date_normalize(df['end_time'])

    return df


# import the two data sets with pandas
annual_data = pd.read_csv("annual_data.csv",
                          parse_dates=['Start Time',
                                       'End Time', 'Date Scheduled'],
                          infer_datetime_format=True,
                          dayfirst=True,
                          cache_dates=True,
                          dtype={
                              'First Name': 'string',
                              'Last Name': 'string',
                              'Phone': 'string',
                              'Type': 'string',
                              'Calendar': 'string',
                              'Appoinment Price': 'float',
                              'Paid?': 'string',
                              'Amount Paid Online': 'float',
                              'Certificate Code': 'string',
                              'Notes': 'string',
                              'Label': 'string',
                              'Scheduled By': 'string',
                              'Name of the company you work for | Име на фирмата, в която работите': 'string',
                              'Preferred platforms | Предпочитани платформи': 'string',
                              'Appointment ID': 'string'},
                          converters={
                              'Email': to_lower, 'Work email | Служебен имейл': to_lower},
                          ).fillna(np.nan).replace([np.nan], [None])

# create a dataframe for the current month only using the last row from the annual build_report as YEAR-MONT value
last_row_value = (annual_data['Start Time'].tail(1).dt.to_period('M'))
monthly_data = annual_data.loc[annual_data['Start Time'].dt.to_period('M').isin(last_row_value)].reset_index(drop=True)


# create flags dictionary for references
flags_dict = {
    'flag_number': [1, 2, 3, 4, 5, 6, 7, 8, 9],
    'flag_note':
        ['names issue',
         'work_mail issue',
         'pvt_mail issue',
         'missing both pvt_email and work_email',
         'missing phone number',
         'phone number issue',
         'missing short_type',
         'pvt email equal to work email',
         'number of the training left less than 1'
         ]
}

limitations = pd.read_csv("limitations.csv",
                          names=['company', 'c_per_emp', 'c_per_month', 'prepaid',
                                 'starts', 'ends', 'note', 'bgn_per_hour', 'is_valid'],
                          usecols=[0, 1, 2, 3, 4, 5, 7, 8, 9],
                          parse_dates=['starts', 'ends'],
                          infer_datetime_format=True,
                          dayfirst=True,
                          cache_dates=True,
                          skiprows=[0],
                          index_col=False,
                          keep_default_na=False,
                          )

# transform the imported .csv files
flags_data = pd.DataFrame(flags_dict)
monthly_data = data_transformation(monthly_data, limitations)
annual_data = data_transformation(annual_data, limitations)

# calculate contract days
limitations['contract_duration'] = dates_diff(limitations['starts'], limitations['ends'])

# normalize date and datetime
limitations['ends'] = date_normalize(limitations['ends'])
limitations['starts'] = datetime_normalize(limitations['starts'])

# '''
# separate and select only needed columns for new pd sets
# '''

new_monthly_data_df = monthly_data[[
    'training_datetime',
    'dayname',
    'month',
    'year',
    'nickname',
    'employee_names',
    'company',
    'trainer',
    'short_type',
    'bgn_per_hour',
    'concat_emp_company',
    'total_per_emp',
    'training_left',
    'returns_or_not',
    'flags',
    'phone',
    'pvt_email',
    'work_email',
    'scheduled_date',
    'type',
    'emp_names_input',
    'is_valid'
]]

new_full_data_df = annual_data[[
    'training_datetime',
    'dayname',
    'month',
    'year',
    'nickname',
    'employee_names',
    'company',
    'trainer',
    'short_type',
    'bgn_per_hour',
    'concat_emp_company',
    'total_per_emp',
    'training_left',
    'returns_or_not',
    'flags',
    'phone',
    'pvt_email',
    'work_email',
    'scheduled_date',
    'type',
    'emp_names_input',
    'is_valid'
]]


total_trainings_df = new_monthly_data_df[[
    'concat_emp_company',
    'type',
    'company',
    'nickname',
    'training_datetime',
    'employee_names',
    'work_email',
    'trainer',
    'short_type'
]]

# add column for the number of trainings left based on the company contract and used trainings by the employee
total_trainings_df = pd.merge(total_trainings_df,
                              new_full_data_df[['concat_emp_company', 'training_datetime', 'trainings_left']], on=[
                                   'concat_emp_company', 'training_datetime'], how='inner')

total_trainings_df = total_trainings_func(total_trainings_df)

total_trainings_df = total_trainings_df[[
    'concat_emp_company',
    'type',
    'company',
    'nickname',
    'training_datetime',
    'employee_names',
    'work_email',
    'trainer',
    'trainings_left',
    'short_type',
    'status',
    'language'
]]



report_trainers = total_trainings_df[[
    'type',
    'trainer',
    'company',
    'employee_names',
    'training_datetime',
    'short_type',
    'status'
]]


limitations['starts'] = pd.to_datetime(
    limitations['starts'], format='%d-%b-%Y %H:%M:%S')  # https://strftime.org/

# '''
# create lists with dfs for later use in to_xlsx sheets
# '''




raw_report_list = ['new_monthly_data',
                   'new_annual_data', 'limitations', 'flags_data']
general_report_list = ['total_trainings', 'report_trainers']

# list the companies with is_valid == 1
company_report_list = [comp for comp in new_monthly_data_df[(
        new_monthly_data_df['is_valid'] == 1)]['company'].unique()]


# list the companies with is_valid == 0
company_report_list_other = [comp_other for comp_other in new_monthly_data_df[(
        new_monthly_data_df['is_valid'] == 0)]['company'].unique()]

# list the trainers names
trainers_report_list = [psy for psy in monthly_data['trainer'].unique()]


# preparing the new df by company (is_valid == 1) in scope + (is_valid == 0) out of the project scope
with pd.ExcelWriter(
        'HourSpacePSY_COMPANIES.xlsx', engine='xlsxwriter') as ew:
    # add worksheets for the consultancies in scope
    for company_name in company_report_list:
        company_filter = (new_monthly_data_df['company'] == company_name)
        new_df = new_monthly_data_df.loc[company_filter, 'nickname'].value_counts().reset_index(name='count')
        new_df.loc[-1, 'total'] = new_df['count'].sum()
        company_df = new_df
        company_df.to_excel(ew, sheet_name=company_name, header=['Employee ID', 'Bookings', 'Total'], index=False)

    # add worksheets for the consultancies out of scope
    for company_name in company_report_list_other:
        company_filter = (new_monthly_data_df['company'] == company_name)
        new_df = new_monthly_data_df.loc[company_filter, 'nickname'].value_counts().reset_index(name='count')
        new_df.loc[-1, 'total'] = new_df['count'].sum()
        company_df = new_df
        company_df.to_excel(ew, sheet_name=company_name, header=['Employee ID', 'Bookings', 'Total'],
                            index=False
                            )


# preparing the new df by therapist
with pd.ExcelWriter(
        'HourSpacePSY_THERAPISTS.xlsx', engine='xlsxwriter') as ew:
    for trainer in trainers_report_list:
        new_df = new_monthly_data_df[(new_monthly_data_df['trainer'] == trainer)][
            ['company', 'training_datetime', 'employee_names', 'short_type', 'bgn_per_hour']]
        # add two columns for totals
        new_df.loc[-1, 'total_consultancies'] = new_df['training_datetime'].count()
        new_df.loc[-1, 'total_pay'] = str(new_df['bgn_per_hour'].sum()) + ".лв"

        company_df = new_df
        company_df.to_excel(ew, sheet_name=trainer,
                            header=['Компания', 'Време на тренинга', 'Имена на служител', 'Вид', 'Лв на час',
                                    'Общо тренинги', 'Общо възнаграждение']  # , index_label='Total'
                            , index=False
                            )


# EXPORTING THE DFs to different sheets
# Create a Pandas Excel writer using XlsxWriter as the engine.
with pd.ExcelWriter(
        'HourSpacePSY_REPORTS_GENERAL.xlsx', engine='xlsxwriter') as ew:
    for df_name in general_report_list:
        df = locals()[df_name]
        df.to_excel(ew, sheet_name=df_name, index=False)

    for df_name in raw_report_list:
        df = locals()[df_name]
        df.to_excel(ew, sheet_name=df_name, index=False)

    month_describe = pd.DataFrame(new_monthly_data_df.describe())
    annual_describe = pd.DataFrame(new_full_data_df.describe())
    month_describe.to_excel(ew, sheet_name='month_describe')
    annual_describe.to_excel(ew, sheet_name='annual_describe')


# TODO export df for monthly stats by company, therapist

with pd.ExcelWriter('HourSpacePSY_REPORTS_STATS_MONTH.xlsx', engine='xlsxwriter') as ew:

    # add sheet for statistical monthly data by Employee
    filter_for_actual_contracts = (new_monthly_data_df['is_valid'] == 1)

    month_stats_emp = new_monthly_data_df[['nickname', 'company']]\
        .value_counts()\
        .reset_index(name='Consultations per Employee')\
        .sort_values(by='nickname')
    month_stats_emp.loc[-1, 'total'] = month_stats_emp['Consultations per Employee'].sum()
    month_stats_emp.to_excel(ew, sheet_name='by employee',
                             index=False,
                             header=['EmployeeID', 'Company', 'Consultancies', 'Total'],
                             freeze_panes=(1, 4))

    # add sheet for statistical monthly data by Company
    month_stats_comp = new_monthly_data_df[['company', 'nickname', 'concat_emp_company']]

    month_stats_comp_pivot = pd.pivot_table(month_stats_comp,
                                            index=['company', 'nickname'],
                                            values=['concat_emp_company'],
                                            aggfunc='count',
                                            margins=True,
                                            margins_name="Total",
                                            fill_value=0,
                                            sort=True
                                            )

    month_stats_comp_pivot.to_excel(ew, sheet_name='by company',
                                    index=['company'],
                                    index_label=['Company', 'EmployeeID'],
                                    header=['Consultancies'],
                                    freeze_panes=(1, 3)
                                    )

    # add sheet for statistical monthly data by Therapist
    month_stats_ther = new_monthly_data_df[['trainer', 'nickname', 'concat_emp_company']]

    month_stats_ther = pd.pivot_table(month_stats_ther,
                                      index=['trainer', 'nickname'],
                                      values=['concat_emp_company'],
                                      aggfunc='count',
                                      margins=True,
                                      margins_name="Total",
                                      fill_value=0,
                                      sort=True
                                      )

    month_stats_ther.to_excel(ew, sheet_name='by therapist',
                              index=['trainer', 'nickname'],
                              index_label=['Therapist', 'EmployeeID'],
                              header=['Consultancies'],
                              freeze_panes=(1, 3)
                              )

    # add sheet for statistical monthly data for Copmany's total
    month_stats_comp_total = new_monthly_data_df[['company', 'concat_emp_company', 'bgn_per_hour']]

    month_stats_comp_total = pd.pivot_table(month_stats_comp_total,
                                            index=['company'],
                                            values=['concat_emp_company', 'bgn_per_hour'],
                                            aggfunc={'concat_emp_company': 'count', 'bgn_per_hour': sum},
                                            margins=True,
                                            margins_name="Total",
                                            fill_value=0,
                                            sort=True
                                            )

    month_stats_comp_total.to_excel(ew, sheet_name='company_total',
                                    index=['company'],
                                    index_label='Company',
                                    header=['BGN', 'Total Consultancies'],
                                    freeze_panes=(1, 3)
                                    )

    # add sheet for statistical monthly data for Therapist's total

    month_stats_ther_total = new_monthly_data_df[['trainer', 'bgn_per_hour', 'is_valid']]

    month_stats_ther_total = pd.pivot_table(month_stats_ther_total,
                                            index=['trainer'],
                                            values=['trainer', 'bgn_per_hour'],
                                            aggfunc={'trainer': 'count', 'bgn_per_hour': sum},
                                            margins=True,
                                            margins_name="Total",
                                            fill_value=0,
                                            sort=True
                                            )

    month_stats_ther_total.to_excel(ew, sheet_name='trainer_total',
                                    index=['trainer'],
                                    index_label='Trainer',
                                    header=['BGN', 'Total Consultancies'],
                                    freeze_panes=(1, 3)
                                    )


# TODO export df for annual stats by company, therapist

with pd.ExcelWriter('HourSpacePSY_REPORTS_STATS_ANNUAL.xlsx', engine='xlsxwriter') as ew:

    # add sheet for statistical annual data by Company with percentage
    filter_for_actual_contracts_y = (new_full_data_df['is_valid'] == 1)

    annual_stats_general1 = new_full_data_df[['company']].value_counts().reset_index()  # .value_counts(normalize=True)*100
    annual_stats_general2 = (new_full_data_df[['company']].value_counts(normalize=True).mul(100).round(2).astype(str) + '%').reset_index()

    annual_stats_general = pd.merge(
        left=annual_stats_general1,
        right=annual_stats_general2,
        left_on='company',
        right_on='company',
        how='left')

    annual_stats_general.to_excel(ew, sheet_name='company_general',
                                  index=False,
                                  header=['Company', 'Total Consultancies', '%Percentage'],
                                  freeze_panes=(1, 3)
                                  )
    # add sheet for statistical annual data by Company and Year
    annual_stats_by_year = new_full_data_df[['company', 'year']]

    # annual_stats_by_year = pd.pivot_table(annual_stats_by_year,
    #                                       index=['company'],
    #                                       columns=['year'],
    #                                       values=['company'],
    #                                       # aggfunc='size',
    #                                       aggfunc={'company': 'count'},
    #                                       # margins=True,
    #                                       # margins_name="Total",
    #                                       fill_value=0,
    #                                       sort=True
    #                                       )
    annual_stats_by_year = annual_stats_by_year.value_counts(['company', 'year']).unstack()  # .fillna(0)
    annual_stats_by_year['Total'] = annual_stats_by_year.agg("sum", axis='columns')

    annual_stats_by_year.loc[-1, 'Grand Total'] = annual_stats_by_year['Total'].sum()

    annual_stats_by_year.to_excel(ew, sheet_name='by_year',
                                  index=['company'],
                                  index_label=['Company'],
                                  # header=['Year', '0'],
                                  freeze_panes=(1, 6)
                                  )

    # add sheet for statistical annual data by Unique EmployeeIDs with percentage
    annual_stats_unique_emp0 = new_full_data_df[['company', 'nickname']].value_counts().to_frame().reset_index()

    annual_stats_unique_emp1 = annual_stats_unique_emp0[['company']].value_counts().to_frame().reset_index()
    annual_stats_unique_emp2 = (annual_stats_unique_emp0[['company']].value_counts(normalize=True).mul(100).round(2).astype(str) + '%').reset_index()

    annual_stats_unique_emp = pd.merge(
        left=annual_stats_unique_emp1,
        right=annual_stats_unique_emp2,
        left_on='company',
        right_on='company',
        how='left')

    annual_stats_unique_emp.to_excel(ew, sheet_name='company_unique_emp',
                                     index=False,
                                     header=['Company', 'Total Consultancies', '%Percentage'],
                                     freeze_panes=(1, 3)
                                     )

    # add sheet for statistical annual data by Employee, by Year
    annual_stats_by_emp_by_year = new_full_data_df[['company', 'nickname', 'year']]

    annual_stats_by_emp_by_year = annual_stats_by_emp_by_year.value_counts(['company', 'nickname', 'year']).unstack()  # .fillna(0)
    annual_stats_by_emp_by_year['Total'] = annual_stats_by_emp_by_year.agg("sum", axis='columns')

    annual_stats_by_emp_by_year.loc[-1, 'Grand Total'] = annual_stats_by_year['Total'].sum()

    annual_stats_by_emp_by_year.to_excel(ew, sheet_name='by_emp_by_year',
                                         index=['company', 'nickname'],
                                         index_label=['Company', 'EmployeeID'],
                                         # header=['Year', '0'],
                                         freeze_panes=(1, 7)
                                         )
    # add sheet for statistical annual data by Company and Employees with only 1 consultation
    annual_stats_by_emp_with_one_cons = new_full_data_df[(new_full_data_df['returns_or_not'] == 'only one session')].reset_index()
    annual_stats_by_emp_with_one_cons = annual_stats_by_emp_with_one_cons[['company', 'nickname']].value_counts().to_frame()

    annual_stats_by_emp_with_one_cons.loc[-1, 'Grand Total'] = annual_stats_by_emp_with_one_cons[0].sum()

    # plot = annual_stats_by_emp_with_one_cons.plot.pie(y='mass', figsize=(5, 5))

    annual_stats_by_emp_with_one_cons.to_excel(ew, sheet_name='by_emp_with_one_cons',
                                               index=['company', 'nickname'],
                                               index_label=['Company', 'EmployeeID'],
                                               header=['Count', 'Grand Total'],  # 'EmployeeID',
                                               freeze_panes=(1, 4)
                                               )

    # # writer = pd.ExcelWriter('farm_data.xlsx', engine='xlsxwriter')
    # annual_stats_by_emp_by_year.to_excel(ew, sheet_name='Sheet1')
    # workbook = ew.book
    # worksheet = ew.sheets['Sheet1']

    # # # Create a chart object.
    # chart = annual_stats_by_emp_by_year.plot.pie(subplots=True)

    # # chart = workbook.add_chart({'type': 'pie'})

    # # Configure the series of the chart from the dataframe data.
    # chart.add_series({
    #     'values':     '=Sheet1!$A$2:$A$8',
    #     'gap':        2,
    # })

    # # Configure the chart axes.
    # chart.set_y_axis({'major_gridlines': {'visible': False}})

    # # Turn off chart legend. It is on by default in Excel.
    # chart.set_legend({'position': 'none'})
