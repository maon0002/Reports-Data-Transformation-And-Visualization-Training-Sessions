# coding: utf8
import os
import logging
from typing import List
import numpy as np
import pandas as pd
from datetime import datetime
import xlsxwriter
import tkinter
from tkinter import filedialog


class Import:

    @staticmethod
    def show_first_ten_rows(path) -> pd.DataFrame:
        """
        Uses a DataFrame to show first 10 rows from it
        :return:
        """
        df = pd.read_csv(path)
        return df.head(10)

    @staticmethod
    def to_lower(x): return x.lower() if isinstance(x, str) else x

    @staticmethod
    def import_report(path) -> pd.DataFrame:
        # tkinter.Tk().withdraw()  # prevents an empty tkinter window from appearing
        # folder_path = filedialog.askdirectory()

        report_df = pd.read_csv(path,
                                parse_dates=['Start Time',
                                             'End Time', 'Date Scheduled'],
                                dayfirst=False,
                                cache_dates=True,
                                dtype={
                                    'First Name': 'string',
                                    'Last Name': 'string',
                                    'Phone': 'string',
                                    'Type': 'string',
                                    'Calendar': 'string',
                                    'Appointment Price': 'float',
                                    'Paid?': 'string',
                                    'Amount Paid Online': 'float',
                                    'Certificate Code': 'string',
                                    'Notes': 'string',
                                    'Label': 'string',
                                    'Scheduled By': 'string',
                                    'Име на компанията, в която работите | Name of the company you work for  ':
                                        'string',
                                    'Предпочитани платформи | Preferred platforms  ': 'string',
                                    'Appointment ID': 'string'},
                                converters={
                                    'Email': Import.to_lower, 'Служебен имейл | Work email  ': Import.to_lower}, ) \
            .fillna(np.nan).replace([np.nan], [None])
        return report_df

    @staticmethod
    def import_limitations(path):
        # tkinter.Tk().withdraw()  # prevents an empty tkinter window from appearing
        # folder_path = filedialog.askdirectory()
        limitations_df = pd.read_csv(path,
                                     names=['company', 'c_per_emp', 'c_per_month', 'prepaid',
                                            'starts', 'ends', 'note', 'bgn_per_hour', 'is_valid'],
                                     usecols=[0, 1, 2, 3, 4, 5, 7, 8, 9],
                                     parse_dates=['starts', 'ends'],
                                     dayfirst=True,
                                     cache_dates=True,
                                     skiprows=[0],
                                     index_col=False,
                                     keep_default_na=False,
                                     )
        return limitations_df


class Export:
    __export_folder = "exports/"

    @staticmethod
    def df_to_excel(dataframe: pd.DataFrame, name: str, suffix: str) -> None:
        """
        Converts the DateFrame with the data and creates a .xlsx file
        :param suffix:
        :param dataframe:
        :param name:
        :return: nothing
        """
        with pd.ExcelWriter(f"{Export.__export_folder}{name}{suffix}.xlsx", engine='xlsxwriter') as ew:
            dataframe.to_excel(ew, index=False)

    @staticmethod
    # def df_to_csv(dataframe: pd.DataFrame, name: str, datetime_format: str) -> None:
    def df_to_csv(dataframe: pd.DataFrame, name: str) -> None:
        """
        Converts the DateFrame with the data to .csv
        :param suffix:
        :param dataframe: transformed from BaseWebsite._data_dict
        :param name: depends on instance name
        :return: nothing
        """
        dataframe.to_csv(f"{Export.__export_folder}{name}.csv", encoding='utf-8', index=False)

    @staticmethod
    def company_trainer_df_to_xlsx(name: str, suffix: str, col_list: list, df: pd.DataFrame, main_col: str,
                                   filter_col: [str or None],
                                   header_cols: list, df_sel_cols: List[str] = None, aggregation: str = 'sum'):
        with pd.ExcelWriter(f'{Export.__export_folder}{name}{suffix}.xlsx', engine='xlsxwriter') as ew:
            # add worksheets for the consultancies in scope
            for value in col_list:

                if filter_col:
                    filter_value = (df[main_col] == value)
                    new_df = df.loc[filter_value, filter_col].value_counts().reset_index(name='count')
                else:
                    new_df = df[(df[main_col] == value)]

                if df_sel_cols:
                    new_df = new_df[df_sel_cols]

                if aggregation == 'sum':
                    new_df.loc[-1, 'total'] = new_df['count'].sum()
                else:
                    new_df.loc[-1, 'total_trainings'] = new_df[df_sel_cols[0]].count()
                    new_df.loc[-1, 'total_pay'] = str(new_df['bgn_per_hour'].sum()) + ".лв"
                new_df.to_excel(ew, sheet_name=value, header=header_cols, index=False)

    @staticmethod
    def general_reports_to_xlsx(name: str, suffix: str, df_general_list: list, df_raw_list: list, dfs: list,
                                df_mont: pd.DataFrame, df_full: pd.DataFrame):
        # def general_reports_to_xlsx(name: str, suffix: str, df_general_list: list, df_raw_list: list,
        # df_mont: pd.DataFrame, df_full: pd.DataFrame):
        with pd.ExcelWriter(f'{name}{suffix}.xlsx', engine='xlsxwriter') as ew:
            # [print(x) for x in dfs]
            [print(x) for x in dfs]
            [getattr(x, name) for x in dfs]
            # for df_name in df_general_list:
            #     if [x for x in dfs if x.__name__ == df_name]:
            #         df = [x for x in dfs if x.__name__ == df_name][0]
            #     df.to_excel(ew, sheet_name=df_name, index=False)
            #
            # for df_name in df_raw_list:
            #     if [x for x in dfs if x.__name__ == df_name]:
            #         df = [x for x in dfs if x.__name__ == df_name][0]
            #     df.to_excel(ew, sheet_name=df_name, index=False)
            #
            # month_describe = pd.DataFrame(df_mont.describe())
            # annual_describe = pd.DataFrame(df_full.describe())
            # month_describe.to_excel(ew, sheet_name='month_describe')
            # annual_describe.to_excel(ew, sheet_name='annual_describe')


class Update:
    pass


class Create:
    __export_folder = "export/"

    @staticmethod
    def create_media_folder(name: str):
        directory = name
        parent_directory = "export/"
        path = os.path.join(parent_directory, directory)
        os.mkdir(path)
        logging.info("Directory '% s' created" % directory)
