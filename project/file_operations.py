# coding: utf8
import logging
from typing import List
import numpy as np
import pandas as pd
from project.invoices import BaseInvoice
import tkinter
from tkinter import filedialog

from project._collections import Collection


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
    def companies_df_to_excel(name: str, dataframe: pd.DataFrame, column_list: List[str] = None):
        # preparing the new df by company (is_valid == 1) in scope + (is_valid == 0) out of the project scope
        with pd.ExcelWriter(f'{Export.__export_folder}{name}.xlsx', engine='xlsxwriter') as ew:
            # add worksheets for the consultancies in scope
            for value in column_list:
                company_filter = (dataframe['company'] == value)

                # use the dataframe filtering for getting data and pass it to the invoice creator
                invoice_data_df_init = dataframe.loc[company_filter][['nickname', 'type']]
                # invoice_data_df = dataframe.loc[company_filter, 'nickname'].value_counts()
                invoice_data_df = invoice_data_df_init.value_counts()
                rate_per_hour = float(dataframe.loc[company_filter, 'bgn_per_hour'].unique())
                invoice_data_dict = invoice_data_df.to_dict()
                BaseInvoice.create_invoice(value, rate_per_hour, invoice_data_dict)
                new_df = dataframe.loc[company_filter, 'nickname'].value_counts().reset_index(name='count')
                new_df.loc[-1, 'total'] = new_df['count'].sum()
                new_df.to_excel(ew, sheet_name=value, header=['Employee ID', 'Bookings', 'Total'],
                                index=False)

    @staticmethod
    def trainers_df_to_excel(name: str, dataframe: pd.DataFrame, column_list: List[str] = None):
        with pd.ExcelWriter(f'{Export.__export_folder}{name}.xlsx', engine='xlsxwriter') as ew:
            for value in column_list:
                new_df = dataframe[(dataframe['trainer'] == value)][
                    ['company', 'training_datetime', 'employee_names', 'short_type', 'bgn_per_hour']]
                # add two columns for totals
                new_df.loc[-1, 'total_consultancies'] = new_df['training_datetime'].count()
                new_df.loc[-1, 'total_pay'] = str(new_df['bgn_per_hour'].sum()) + ".лв"
                new_df.to_excel(ew, sheet_name=value,
                                header=['Компания', 'Време на тренинга', 'Имена на служител', 'Вид', 'Лв на час',
                                        'Общо тренинги', 'Общо възнаграждение'], index=False)

    @staticmethod
    def generic_df_to_excel(name: str, df_dict):
        with pd.ExcelWriter(f'{Export.__export_folder}{name}.xlsx', engine='xlsxwriter') as ew:
            for item in df_dict.items():
                df = item[1]
                df_name = item[0]
                if df_name in Collection.generic_report_list():
                    df.to_excel(ew, sheet_name=str(df_name), index=False)
            month_describe = pd.DataFrame(df_dict["new_monthly_data_df"].describe())
            annual_describe = pd.DataFrame(df_dict["new_full_data_df"].describe())
            month_describe.to_excel(ew, sheet_name='month_describe')
            annual_describe.to_excel(ew, sheet_name='annual_describe')

    @staticmethod
    def stats_mont_df_to_excel(name: str, mont_df: pd.DataFrame):
        with pd.ExcelWriter(f'{Export.__export_folder}{name}.xlsx', engine='xlsxwriter') as ew:
            # add sheet for statistical monthly data by Employee
            filter_for_actual_contracts = (mont_df['is_valid'] == 1)
            month_stats_emp = mont_df[['nickname', 'company']] \
                .value_counts() \
                .reset_index(name='Consultations per Employee') \
                .sort_values(by='nickname')
            month_stats_emp.loc[-1, 'total'] = month_stats_emp['Consultations per Employee'].sum()
            month_stats_emp.to_excel(ew, sheet_name='by employee',
                                     index=False,
                                     header=['EmployeeID', 'Company', 'Consultancies', 'Total'],
                                     freeze_panes=(1, 4))
            # add sheet for statistical monthly data by Company
            month_stats_comp = mont_df[['company', 'nickname', 'concat_emp_company']]
            month_stats_comp_pivot = pd.pivot_table(month_stats_comp,
                                                    index=['company', 'nickname'],
                                                    values=['concat_emp_company'],
                                                    aggfunc='count',
                                                    margins=True,
                                                    margins_name="Total",
                                                    fill_value=0,
                                                    sort=True)
            month_stats_comp_pivot.to_excel(ew, sheet_name='by company',
                                            index=['company'],
                                            index_label=['Company', 'EmployeeID'],
                                            header=['Consultancies'],
                                            freeze_panes=(1, 3))
            # add sheet for statistical monthly data by Trainer
            month_stats_trainer = mont_df[['trainer', 'nickname', 'concat_emp_company']]
            month_stats_trainer = pd.pivot_table(month_stats_trainer,
                                                 index=['trainer', 'nickname'],
                                                 values=['concat_emp_company'],
                                                 aggfunc='count',
                                                 margins=True,
                                                 margins_name="Total",
                                                 fill_value=0,
                                                 sort=True)
            month_stats_trainer.to_excel(ew, sheet_name='by Trainer',
                                         index=['trainer', 'nickname'],
                                         index_label=['Trainer', 'EmployeeID'],
                                         header=['Consultancies'],
                                         freeze_panes=(1, 3))
            # add sheet for statistical monthly data for Copmany's total
            month_stats_comp_total = mont_df[['company', 'concat_emp_company', 'bgn_per_hour']]
            month_stats_comp_total = pd.pivot_table(month_stats_comp_total,
                                                    index=['company'],
                                                    values=['concat_emp_company', 'bgn_per_hour'],
                                                    aggfunc={'concat_emp_company': 'count', 'bgn_per_hour': sum},
                                                    margins=True,
                                                    margins_name="Total",
                                                    fill_value=0,
                                                    sort=True)
            month_stats_comp_total.to_excel(ew, sheet_name='company_total',
                                            index=['company'],
                                            index_label='Company',
                                            header=['BGN', 'Total Consultancies'],
                                            freeze_panes=(1, 3))
            # add sheet for statistical monthly data for Trainer's total
            month_stats_trainer__total = mont_df[['trainer', 'bgn_per_hour', 'is_valid']]
            month_stats_trainer__total = pd.pivot_table(month_stats_trainer__total,
                                                        index=['trainer'],
                                                        values=['trainer', 'bgn_per_hour'],
                                                        aggfunc={'trainer': 'count', 'bgn_per_hour': sum},
                                                        margins=True,
                                                        margins_name="Total",
                                                        fill_value=0,
                                                        sort=True)
            month_stats_trainer__total.to_excel(ew, sheet_name='trainer_total',
                                                index=['trainer'],
                                                index_label='Trainer',
                                                header=['BGN', 'Total Consultancies'],
                                                freeze_panes=(1, 3))

    @staticmethod
    def stats_full_df_to_excel(name: str, full_df: pd.DataFrame):
        with pd.ExcelWriter(f'{Export.__export_folder}{name}.xlsx', engine='xlsxwriter') as ew:
            # add sheet for statistical annual data by Company with percentage
            filter_for_actual_contracts_y = (full_df['is_valid'] == 1)
            annual_stats_general1 = full_df[['company']].value_counts().reset_index()
            annual_stats_general2 = (full_df[['company']].
                                     value_counts(normalize=True).mul(100).round(2).astype(str) + '%').reset_index()
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
            annual_stats_by_year = full_df[['company', 'year']]
            annual_stats_by_year = annual_stats_by_year.value_counts(['company', 'year']).unstack()  # .fillna(0)
            annual_stats_by_year['Total'] = annual_stats_by_year.agg("sum", axis='columns')
            annual_stats_by_year.loc[-1, 'Grand Total'] = annual_stats_by_year['Total'].sum()
            annual_stats_by_year.to_excel(ew, sheet_name='by_year',
                                          index=['company'],
                                          index_label=['Company'],
                                          freeze_panes=(1, 6))
            # add sheet for statistical annual data by Unique EmployeeIDs with percentage
            annual_stats_unique_emp0 = full_df[['company', 'nickname']].value_counts().to_frame().reset_index()
            annual_stats_unique_emp1 = annual_stats_unique_emp0[['company']].value_counts().to_frame().reset_index()
            annual_stats_unique_emp2 = (annual_stats_unique_emp0[['company']].
                                        value_counts(normalize=True).mul(100).round(2).astype(str) + '%').reset_index()
            annual_stats_unique_emp = pd.merge(
                left=annual_stats_unique_emp1,
                right=annual_stats_unique_emp2,
                left_on='company',
                right_on='company',
                how='left')
            annual_stats_unique_emp.to_excel(ew, sheet_name='company_unique_emp',
                                             index=False,
                                             header=['Company', 'Total Consultancies', '%Percentage'],
                                             freeze_panes=(1, 3))
            # add sheet for statistical annual data by Employee, by Year
            annual_stats_by_emp_by_year = full_df[['company', 'nickname', 'year']]
            annual_stats_by_emp_by_year = annual_stats_by_emp_by_year.value_counts(['company', 'nickname', 'year']) \
                .unstack()
            annual_stats_by_emp_by_year['Total'] = annual_stats_by_emp_by_year.agg("sum", axis='columns')
            annual_stats_by_emp_by_year.loc[-1, 'Grand Total'] = annual_stats_by_year['Total'].sum()
            annual_stats_by_emp_by_year.to_excel(ew, sheet_name='by_emp_by_year',
                                                 index=['company', 'nickname'],
                                                 index_label=['Company', 'EmployeeID'],
                                                 freeze_panes=(1, 7))
            # add sheet for statistical annual data by Company and Employees with only 1 consultation
            annual_stats_by_emp_with_one_cons = full_df[(full_df['returns_or_not'] == 'only one session')].reset_index()
            annual_stats_by_emp_with_one_cons = pd.DataFrame(
                annual_stats_by_emp_with_one_cons[['company', 'nickname']].value_counts().to_frame())
            annual_stats_by_emp_with_one_cons.loc[-1, 'Grand Total'] = annual_stats_by_emp_with_one_cons['count'].sum()
            # plot = annual_stats_by_emp_with_one_cons.plot.pie(y='mass', figsize=(5, 5))
            annual_stats_by_emp_with_one_cons.to_excel(ew, sheet_name='by_emp_with_one_cons',
                                                       index=['company', 'nickname'],
                                                       index_label=['Company', 'EmployeeID'],
                                                       header=['Count', 'Grand Total'],
                                                       freeze_panes=(1, 4))
