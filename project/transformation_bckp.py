# coding: utf8
from typing import List
import pandas as pd
import logging

from project.dataframes import BaseDataframe
from project.file_operations import Import, Transform, Export
from project.collections import Collection

logging.basicConfig(filename='info.log', encoding='utf-8',
                    level=logging.INFO,
                    format=u'%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%d-%m-%Y %H:%M:%S'
                    )


class Transformation:
    _date_default_format = "%d-%b-%Y"
    _datetime_default_format = "%d-%b-%Y %H:%M:%S"
    _datetime_final_format = "%Y-%m-%d %H:%M:%S"
    """

    """
    # create flags dictionary for references
    _flags_dict = {
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
             'number of the trainings left less than 1'
             ]}

    def __init__(self, name):
        self.name = name
        self.dataframes_list: List[BaseDataframe] = []

    @staticmethod
    def trim_all_columns(df):
        """
        Trim whitespace from ends of each value across all series in dataframe
        """

        def trim_strings(x): return x.strip() if isinstance(x, str) else x

        return df.applymap(trim_strings)

    @staticmethod
    def unwanted_chars(df):
        """
        Get rid of single quotes and apostrophes
        """
        df = df.replace("'|  |`", "", regex=True)
        return df

    @staticmethod
    def nickname(df: pd.DataFrame) -> pd.DataFrame:
        """
        The nickname must must have the following pattern:
        2 letters from the first name,
        3 letters from the last name and
        3 letters from the email (pvt or work w/o only alpha chars)
        """
        # adding column with both Fname and Lname
        df['employee_names'] = df['first_name'] + " " + df['last_name']

        # adding nickname column
        df['nickname'] = df['first_name'].str[0:2].str.upper() + df['last_name'].str[1:4].str.upper()

        # adding flag and 'X's at the end if the len < 5
        df.loc[df.nickname.str.len() < 5, 'flags'] += '1,'
        df.loc[df.nickname.str.len() < 5, 'nickname'] = df['nickname'].apply(lambda x: x + 'X' * (5 - len(x)))

        # check for email columns and take last part of the nickname
        # add flags for the cases
        work_email_validation_filter = (df['work_email'].str.contains('@', na=False))
        pvt_email_validation_filter = (df['pvt_email'].str.contains('@', na=False))

        df.loc[~work_email_validation_filter, 'flags'] += "2,"
        df.loc[~pvt_email_validation_filter, 'flags'] += "3,"

        df.loc[~pvt_email_validation_filter & ~work_email_validation_filter, 'flags'] += "4,"

        # takes value subtraction from the work or pvt email:
        email_validation_filter = (df['work_email'].str.contains('@', na=False))
        df.loc[email_validation_filter, 'nickname'] += \
            df['work_email'].str.replace("[^A-Za-z]+", "", regex=True).str[1:4].str.upper()

        df.loc[(df['work_email'].str.len() < 1) &
               (~df['pvt_email'].str.len() < 1) &
               (df['pvt_email'].str.contains('@')), 'nickname'] += \
            df['pvt_email'].str.replace("[^A-Za-z]+", "", regex=True).str[1:4].str.upper()

        df.loc[(df['work_email'].str.len() < 1) & (df['pvt_email'].str.len() < 1), 'nickname'] = \
            df['nickname'].apply(lambda x: x + 'X' * (8 - len(x)))

        df.loc[df['nickname'].str.len() < 8, 'nickname'] = df['nickname'].apply(lambda x: x + 'X' * (8 - len(x)))

        # adding column with both Fname and Lname
        df['employee_names'] = df['employee_names'].str.title()

        return df

    @staticmethod
    def company_subtraction(df: pd.DataFrame) -> pd.DataFrame:
        """
        1.Check if the meeting was held online or in-person and mark in a separate column
        2.Iterate through the 'type' column and extracts only the company name
        """

        df.loc[df['type'].str.contains('person|живо', regex=True), 'short_type'] = "На живо"
        df.loc[df['type'].str.contains('nline|нлайн', regex=True), 'short_type'] = "Онлайн"
        df.loc[~df['type'].str.contains('nline|нлайн|person|живо', regex=True), 'flags'] += "7,"

        df.loc[~df['type'].str.len() < 1, 'company'] = df['type'].str.split("[:|/]").str[0].str.upper().str.strip()
        return df

    @staticmethod
    def training_per_emp(df: pd.DataFrame) -> pd.DataFrame:
        """
        Add three additional columns for:
            concatenation of employee|company,
            training calculation total by the concat values,
            either the employee have one or more than one consultation
        """
        # concatenate nickname+company for later count and check if they are equal nicknames from different companies
        df.loc[~df['nickname'].str.len() < 1, 'concat_emp_company'] = df['nickname'].str[:] + "|" + df['company'].str[:]

        # count the frequency of occurrence for every  value (employee)
        df['total_per_emp'] = df.groupby('concat_emp_company')['concat_emp_company'].transform('count')

        # comment if the employee have one or more than one trainings
        df.loc[df['total_per_emp'].astype(int) == 1, 'returns_or_not'] = 'only one session'

        df.loc[df['total_per_emp'].astype(int) > 1, 'returns_or_not'] = 'more then one session'
        return df

    @staticmethod
    def phone_validation(df: pd.DataFrame) -> pd.DataFrame:
        df.loc[df['phone'].str.len() < 1, 'flags'] += "5,"
        df.loc[df['phone'].str.contains('[^\d\+]', regex=True, na=False), 'flags'] += "6,"
        df.loc[df['phone'].str.len().between(1, 8), 'flags'] += "6,"
        return df

    @staticmethod
    def trainer(df: pd.DataFrame) -> pd.DataFrame:
        """
        Subtract only the non cyrillic names and put them in a separate column
        """
        df.loc[~df['calendar'].isnull(), 'trainer'] = df['calendar'].str.split('|').str[0].str.strip().str.title()
        return df

    @staticmethod
    def active_contracts(df: pd.DataFrame) -> pd.DataFrame:
        """
        Count the trainings of the company employees based on the training
        date if there's an active contract'
        """
        # CONCAT the lines for which the training date is between the contract's start and end date
        df.loc[(df['start_time'] >= df['starts']) & (df['start_time'] <= df['ends']), 'concat_count'] = \
            df['company'] + "|" + df['nickname']

        # count the company/employees total trainings for the company active period
        df.loc[(df['start_time'] >= df['starts']) & (df['start_time'] <= df['ends']), 'active_trainings_per_client'] = \
            df.groupby('concat_count')['concat_count'].transform('count')

        # calculate the number of trainings that can be used
        df.loc[(df['start_time'] >= df['starts']) &
               (df['start_time'] <= df['ends']) &
               (df['c_per_emp'].between(1, 9998)), 'trainings_left'] = \
            df['c_per_emp'] - df['active_trainings_per_client']

        # add flag for employees which have less than 1 training left
        df.loc[(~df['trainings_left'].isna()) & (df['trainings_left'] < 2), 'flags'] += '9,'
        return df

    @staticmethod
    def datetime_normalize(df: pd.DataFrame) -> pd.DataFrame:
        df = df.apply(lambda x: x.strftime(Transformation._datetime_default_format))
        return df

    @staticmethod
    def date_normalize(df: pd.DataFrame) -> pd.DataFrame:
        df = pd.to_datetime(df).dt.date
        df = df.apply(lambda x: x.strftime(Transformation._date_default_format))
        return df

    @staticmethod
    def dates_diff(ser1: pd.Series, ser2: pd.Series) -> pd.Series:
        diff = ser2 - ser1
        return diff

    @staticmethod
    def total_trainings_func(df_mont: pd.DataFrame, df_full: pd.DataFrame) -> pd.DataFrame:
        # get and transform the trainings data on a total level
        trainings_column_list_init = Collection.trainings_columns()[0]
        df = df_mont[trainings_column_list_init]

        # add column for the number of trainings left based on the company contract and used trainings by the employee
        df = pd.merge(df, df_full[['concat_emp_company', 'training_datetime', 'trainings_left']],
                      on=['concat_emp_company', 'training_datetime'], how='inner')

        # insert two additional columns with default values
        df.insert(10, "language", 'Български', allow_duplicates=False)
        df.insert(10, "status", 'Проведен', allow_duplicates=False)
        trainings_column_list_final = Collection.trainings_columns()[1]
        df = df[trainings_column_list_final]
        return df

    @staticmethod
    def limitations_func(limitations_df: pd.DataFrame):
        # calculate contract days
        limitations_df['contract_duration'] = Transformation.dates_diff(limitations_df['starts'],
                                                                        limitations_df['ends'])
        # normalize date and datetime
        limitations_df['ends'] = Transformation.date_normalize(limitations_df['ends'])
        limitations_df['starts'] = Transformation.datetime_normalize(limitations_df['starts'])
        limitations_df['starts'] = pd.to_datetime(limitations_df['starts'], format='%d-%b-%Y %H:%M:%S')
        return limitations_df

    @staticmethod
    def generic_reports_preparation(df_name):
        df = locals()[df_name]
        return df

    def exporting(self, full_raw_report_df,
                  monthly_raw_report_df,
                  new_full_data_df,
                  new_monthly_data_df,
                  total_trainings_df):
        # export the dfs to .csv/.xlsx

        Export.df_to_csv(full_raw_report_df, self.name, "_full_raw_report")
        Export.df_to_csv(monthly_raw_report_df, self.name, "_monthly_raw_report")
        Export.df_to_csv(new_full_data_df, self.name, "_NEW_FULL_REPORT")
        Export.df_to_csv(monthly_raw_report_df, self.name, "_NEW_MONTHLY_REPORT")

        Export.company_trainer_df_to_xlsx(self.name, "_COMPANIES",
                                          Collection.company_report_list(new_monthly_data_df),
                                          new_monthly_data_df,
                                          'company', 'nickname', ['Employee ID', 'Bookings', 'Total'])
        Export.company_trainer_df_to_xlsx(self.name, "_COMPANIES_OTHER",
                                          Collection.company_report_list_other(new_monthly_data_df),
                                          new_monthly_data_df,
                                          'company', 'nickname', ['Employee ID', 'Bookings', 'Total'])
        Export.company_trainer_df_to_xlsx(self.name, "_TRAINERS",
                                          Collection.trainers_report_list(new_monthly_data_df),
                                          new_monthly_data_df,
                                          'trainer', None,
                                          ['Компания', 'Време на тренинга', 'Имена на служител', 'Вид', 'Лв на час',
                                           'Общо тренинги', 'Общо възнаграждение'],
                                          ['company', 'training_datetime', 'employee_names', 'short_type',
                                           'bgn_per_hour'],
                                          "count")

        Export.general_reports_to_xlsx(self.name, "_REPORTS_GENERAL", Collection.general_report_list(),
                                       Collection.raw_report_list(), self.dataframes_list, new_monthly_data_df,
                                       new_full_data_df)

    def main(self, df: pd.DataFrame, limitations_df: pd.DataFrame) -> List[pd.DataFrame]:
        """
        Transform the monthly and the annual df columns
        """
        flags_data = pd.DataFrame(Collection.flags_dict())
        limitations = self.limitations_func(limitations_df)
        self.dataframes_list.append(BaseDataframe("new_limitations", limitations))
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
            'Име на компанията, в която работите | Name of the company you work for  ': 'company_name',
            'Служебен имейл | Work email  ': 'work_email',
            'Предпочитани платформи | Preferred platforms  ': 'preferred_platforms',
            'Appointment ID': 'appointment_id'
        })

        # add column for issues
        df['flags'] = ""
        df['emp_names_input'] = '|' + df['f_name'] + '|' + df['l_name'] + '|'

        # rough cleaning of the columns data
        df = self.unwanted_chars(df)
        df = self.trim_all_columns(df)

        # transliteration of cyrillic to latin chars
        df["first_name"] = Collection.transliterate_bg_to_en(df, "f_name", "first_name")
        df["last_name"] = Collection.transliterate_bg_to_en(df, "l_name", "last_name")

        # using first, last name and email parts
        df = self.nickname(df)

        # get only the company name and if the training was IN PERSON/LIVE or ONLINE
        df = self.company_subtraction(df)

        # merge/vlookup the columns from limitations to the monthly/annual df
        df = pd.merge(
            left=df,
            right=limitations,
            left_on='company',
            right_on='company',
            how='left')

        # add columns for counting unique emp|company values w/ totals
        df = self.training_per_emp(df)

        # check phone values
        df = self.phone_validation(df)

        # substring trainers from calendar via regex.
        df = self.trainer(df)

        # check if the training date is between the dates when company contract starts and ends
        df = self.active_contracts(df)

        # create a Month name column
        df["month"] = [pd.Timestamp(x).month_name() for x in df["start_time"]]

        # create a Year name column
        df["year"] = [pd.Timestamp(x).year for x in df["start_time"]]

        # add column with the name of the day when training was take place
        df['dayname'] = df['start_time'].dt.day_name()

        # reformat start_time
        df['training_datetime'] = self.datetime_normalize(df['start_time'])

        # change from datetime to date only
        df['scheduled_date'] = self.date_normalize(df['scheduled_on'])
        df['training_end'] = self.date_normalize(df['end_time'])

        # define/make the raw reports and extract them to .csv
        full_raw_report_df = df
        monthly_raw_report_df = Transform.annual_to_monthly_report_df(full_raw_report_df,
                                                                      Transformation._datetime_final_format)

        # separate and select only needed columns for new pd sets
        columns_list = Collection.new_data_columns()

        new_full_data_df = full_raw_report_df[columns_list]
        new_monthly_data_df = monthly_raw_report_df[columns_list]

        total_trainings_df = self.total_trainings_func(new_monthly_data_df, new_full_data_df)
        self.dataframes_list.append(BaseDataframe("raw_report_full", full_raw_report_df))
        self.dataframes_list.append(BaseDataframe("raw_report_mont", monthly_raw_report_df))
        self.dataframes_list.append(BaseDataframe("new_report_full", new_full_data_df))
        self.dataframes_list.append(BaseDataframe("new_report_mont", new_monthly_data_df))
        self.dataframes_list.append(BaseDataframe("total_trainings", total_trainings_df))


        self.exporting(full_raw_report_df,
                       monthly_raw_report_df,
                       new_full_data_df,
                       new_monthly_data_df,
                       total_trainings_df)

        return [full_raw_report_df,
                monthly_raw_report_df,
                new_full_data_df,
                new_monthly_data_df,
                total_trainings_df,
                ]


firm = Transformation("ServiceCompanyEOOD")
Transformation.main(firm,
                    Import.import_report('imports/schedule2023-04-18.csv'),
                    Import.import_limitations('imports/limitations.csv'))
