# coding: utf8
from typing import List
import pandas as pd
import logging
from project._collections import Collection


class BaseDataframe:
    """Class used to handle dataframe operations

        Attributes
        ----------
        No attributes

        Methods
        -------
        rename_original_report_columns(df: pd.DataFrame) -> pd.DataFrame:
            Renames the columns from the imported general report

        trim_all_columns(df: pd.DataFrame) -> pd.DataFrame:
            Trim leading and lagging white spaces(if any)

        unwanted_chars(df: pd.DataFrame) -> pd.DataFrame:
            Get rid of single quotes and apostrophes

        def transliterate_bg_to_en(df: pd.DataFrame, column: str, new_column: str) -> pd.Series:
            Uses a dictionary with the transliteration pairs {BG:EN} to
            transliterate pd.Series / columns

        nickname(df: pd.DataFrame) -> pd.DataFrame:
            Validating and transforming the 'name' and 'email' related columns
            and creating an additional one for the purpose of
            attaching a nickname to every employee
            in a not-so-recognisable way

        company_subtraction(df: pd.DataFrame) -> pd.DataFrame:
            Checks if the meeting was held online or in-person and mark in a separate column
            Iterate through the 'type' column and extracts only the company name

        training_per_emp(df: pd.DataFrame) -> pd.DataFrame:
            Add three additional columns for:
            1.concatenation of employee|company,
            2.training calculation total by the concat values,
            3.either the employee have one or more than one consultation

        phone_validation(df: pd.DataFrame) -> pd.DataFrame:
            Add different flags if non standard phone values are present in the series

        trainer(df: pd.DataFrame) -> pd.DataFrame:
            Subtract only the non cyrillic names and put them in a separate column

        active_contracts(df: pd.DataFrame) -> pd.DataFrame:
            Count the trainings of the company employees based on the training
            date if there's an active contract (record is present in the limitations.csv)

        datetime_normalize(df: pd.DataFrame) -> pd.DataFrame:
            Handle the datetime formatting using the format from the _collections
            for the 'training_datetime' and 'starts' columns

        date_normalize(df: pd.DataFrame) -> pd.DataFrame:
            Handle the date formatting using the format from the _collections
            for the 'scheduled_date', 'training_end'  and 'ends' columns

        dates_diff(ser1: pd.Series, ser2: pd.Series) -> pd.Series:
            Adds column 'contract_duration' which is the days difference between
            the start and the end of the contract for a particular company (limitation.csv)

        total_trainings_func(df_mont: pd.DataFrame, df_full: pd.DataFrame) -> List[pd.DataFrame]:
            Transform the training sessions on a total level by company and by trainer

        limitations_func(limitations_df: pd.DataFrame) -> pd.DataFrame:
            Adds three additional columns in the limitations dataframe

    """

    @staticmethod
    def rename_original_report_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        Renames the columns from the imported general report
        with lowercase string with underscores

        :param df: initial dataframe from the imported .csv general report
        :return: same dataframe with changed column names
        """
        df_new_column_names = df.rename(columns={
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
        return df_new_column_names

    @staticmethod
    def trim_all_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        Trim leading and lagging white spaces(if any)
        for all values in the given dataframe

        :param df: the dataframe from the imported .csv general report in a current stage
                (after some additional transformations)
        :return: same dataframe with no blanks in the beginning of the value or at the end
        """

        def trim_strings(x): return x.strip() if isinstance(x, str) else x

        return df.applymap(trim_strings)

    @staticmethod
    def unwanted_chars(df: pd.DataFrame) -> pd.DataFrame:
        """
        Get rid of single quotes and apostrophes
        which can mess the values from 'nickname' series/columns

        :param df: the dataframe from the imported .csv general report in a current stage
                (after some additional transformations)
        :return: same dataframe with no apostrophes or single quotes in values
        """

        df = df.replace("'|  |`", "", regex=True)
        return df

    @staticmethod
    def transliterate_bg_to_en(df: pd.DataFrame, column: str, new_column: str) -> pd.Series:
        """
        Uses a dictionary with the transliteration pairs {BG:EN} to
        transliterate pd.Series / columns

        :param df:
        :param column:
        :param new_column:
        :return:
        """

        bg_en_dict = Collection.transliterate_dict()

        # transliteration char by char
        transliterated_string = ""
        for i in range(len(df[column])):
            string_value = df.loc[i, column].strip()
            if string_value and not string_value.isascii():
                for char in string_value:
                    if char in bg_en_dict.keys():
                        string_value = string_value.replace(
                            char, bg_en_dict[char])
                        transliterated_string = string_value
                    else:
                        transliterated_string = string_value

                df.loc[i, new_column] = transliterated_string.upper()
                transliterated_string = ""
            else:
                df.loc[i, new_column] = df.loc[i, column].upper()
        return df[new_column]

    @staticmethod
    def nickname(df: pd.DataFrame) -> pd.DataFrame:
        """
        Validating and transforming the 'name' and 'email' related columns
        and creating an additional one for the purpose of
        attaching a nickname to every employee
        in a not-so-recognisable way

        The nicknames are build based on the following pattern:
        2 letters from the first name,
        3 letters from the last name and
        3 letters from the email (pvt or work w/o only alpha chars)
        *if for some reason (empty values in the needed columns)
            the length of the nickname is less than 8 chars
            it adds lagging 'X's until the 8th position

        :param df: the dataframe from the imported .csv general report in a current stage
                (after some additional transformations)
        :return: same dataframe with additional column 'nickname' (+ flags if any)
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

        # adding column with both first and last names
        df['employee_names'] = df['employee_names'].str.title()
        return df

    @staticmethod
    def company_subtraction(df: pd.DataFrame) -> pd.DataFrame:
        """
        Checks if the meeting was held online or in-person and mark in a separate column
        Iterate through the 'type' column and extracts only the company name

        :param df: the dataframe from the imported .csv general report in a current stage
                (after some additional transformations)
        :return: the dataframe with two additional columns: 'company' and 'short_type'
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
        1.concatenation of employee|company,
        2.training calculation total by the concat values,
        3.either the employee have one or more than one consultation

        :param df: the dataframe from the imported .csv general report in a current stage
                (after some additional transformations)
        :return: the dataframe with two additional columns: 'concat_emp_company', 'count' and 'returns_or_not'
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
        """
        Add different flags if non standard phone values are present in the series

        :param df: the dataframe from the imported .csv general report in a current stage
                (after some additional transformations)
        :return: the dataframe with added values in the flag column
        """
        df.loc[df['phone'].str.len() < 1, 'flags'] += "5,"
        df.loc[df['phone'].str.contains(r'[^\d\+]', regex=True, na=False), 'flags'] += "6,"
        df.loc[df['phone'].str.len().between(1, 8), 'flags'] += "6,"
        return df

    @staticmethod
    def trainer(df: pd.DataFrame) -> pd.DataFrame:
        """
        Subtract only the non cyrillic names and put them in a separate column

        :param df: the dataframe from the imported .csv general report in a current stage
                (after some additional transformations)
        :return: the dataframe with one additional column: 'trainer'
        """
        df.loc[~df['calendar'].isnull(), 'trainer'] = df['calendar'].str.split('|').str[0].str.strip().str.title()
        return df

    @staticmethod
    def active_contracts(df: pd.DataFrame) -> pd.DataFrame:
        """
        Count the trainings of the company employees based on the training
        date if there's an active contract (record is present in the limitations.csv)

        :param df: the dataframe from the imported .csv general report in a current stage
                (after some additional transformations)
        :return: the dataframe with three additional columns:
                'trainings_left', 'active_trainings_per_client', 'concat_count'
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
        """
        Handle the datetime formatting using the format from the _collections
        for the 'training_datetime' and 'starts' columns

        :param df: a dataframe
        :return: same dataframe with normalized datetime column
        """
        df = df.apply(lambda x: x.strftime(Collection.datetime_default_format()))
        return df

    @staticmethod
    def date_normalize(df: pd.DataFrame) -> pd.DataFrame:
        """
        Handle the date formatting using the format from the _collections
        for the 'scheduled_date', 'training_end'  and 'ends' columns

        :param df: a dataframe
        :return: same dataframe with normalized date column
        """
        df = pd.to_datetime(df).dt.date
        df = df.apply(lambda x: x.strftime(Collection.date_default_format()))
        return df

    @staticmethod
    def dates_diff(ser1: pd.Series, ser2: pd.Series) -> pd.Series:
        """
        Adds column 'contract_duration' which is the days difference between
        the start and the end of the contract for a particular company (limitation.csv)

        :param ser1: limitations_df['starts']
        :param ser2: limitations_df['ends']
        :return: a series with the number of days as int
        """
        diff = ser2 - ser1
        return diff

    @staticmethod
    def total_trainings_func(df_mont: pd.DataFrame, df_full: pd.DataFrame) -> List[pd.DataFrame]:
        """
        Transform the training sessions on a total level by company and by trainer
        adding two additional columns with default values: 'language' and 'status'
        using dataframe series for the both data sets predefined in _collections

        :param df_mont: new_monthly_data_df (a transformed full_raw_report_df)
        :param df_full: new_full_data_df (a transformed monthly_raw_report_df)
        :return: list with two new dataframes
        """
        # get and transform the trainings data on a total level
        # using only the columns from the Collection class
        trainings_column_list_init = Collection.trainings_columns()[0]
        df = df_mont[trainings_column_list_init]

        # add column for the number of trainings left based on the company contract and used trainings by the employee
        df = pd.merge(df, df_full[['concat_emp_company', 'training_datetime', 'trainings_left']],
                      on=['concat_emp_company', 'training_datetime'], how='inner')
        # insert two additional columns with default values
        df.insert(10, "language", 'Български', allow_duplicates=False)
        df.insert(10, "status", 'Проведен', allow_duplicates=False)

        total_trainings_df = df

        # get and transform the trainings data on a total level
        # using only the columns from the Collection class
        trainings_column_list_final = Collection.trainings_columns()[1]

        report_trainers_df = df[trainings_column_list_final]
        return [total_trainings_df, report_trainers_df]

    @staticmethod
    def limitations_func(limitations_df: pd.DataFrame) -> pd.DataFrame:
        """
        Adds three additional columns in the limitations dataframe
        based on limitations.csv import:
            'contract_duration', 'starts', 'ends'

        :param limitations_df:
        :return: the dataframe with the three additional columns
        """
        # calculate contract days
        limitations_df['contract_duration'] = BaseDataframe.dates_diff(limitations_df['starts'],
                                                                       limitations_df['ends'])
        # normalize date and datetime
        limitations_df['ends'] = BaseDataframe.date_normalize(limitations_df['ends'])
        limitations_df['starts'] = BaseDataframe.datetime_normalize(limitations_df['starts'])
        limitations_df['starts'] = pd.to_datetime(limitations_df['starts'], format=Collection.datetime_default_format())
        return limitations_df
