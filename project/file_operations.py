# coding: utf8
import logging
from datetime import datetime
import zipfile
import glob
import shutil
import os
from typing import List
import numpy as np
import pandas as pd
from project.invoices import BaseInvoice
import tkinter
from tkinter import filedialog
from tkinter.filedialog import asksaveasfile
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
    def to_lower(x):
        return x.lower() if isinstance(x, str) else x

    @staticmethod
    def import_report() -> pd.DataFrame:
        expected_columns = ['Start Time', 'End Time', 'First Name', 'Last Name', 'Phone', 'Email', 'Type', 'Calendar',
                            'Appointment Price', 'Paid?', 'Amount Paid Online', 'Certificate Code', 'Notes',
                            'Date Scheduled', 'Label', 'Scheduled By',
                            'Име на компанията, в която работите | Name of the company you work for  ',
                            'Служебен имейл | Work email  ', 'Предпочитани платформи | Preferred platforms  ',
                            'Appointment ID']
        report_df = None
        while True:
            tkinter.Tk().withdraw()  # prevents an empty tkinter window from appearing
            file_path = filedialog.askopenfilename(title='Select your report to import the data',
                                                   filetypes=[("CSV (comma-separated values) file", ".csv")],
                                                   defaultextension=".csv",
                                                   )
            if not file_path:
                print(f"Please choose the .csv file with the General report!")
                continue

            columns_in_selected_file = pd.read_csv(file_path).columns
            column_inconsistency_check = [x for x in zip(expected_columns, columns_in_selected_file) if x[0] != x[1]]

            if column_inconsistency_check:
                print(column_inconsistency_check, sep="\n")
                message = [">>> The following columns are not matching expected order/names: "]
                for items in column_inconsistency_check:
                    expected = items[0]
                    found = items[1]
                    line_to_add = f"The expected column: '{expected}' is not equal to the '{found}' " \
                                  f"column from the chosen file!"
                    message.append(line_to_add)

                print(f"Please check the column's consistency in the chosen file!\n",
                      '\n'.join(message))
                continue
            else:
                report_df = pd.read_csv(file_path,
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
                                            'Email': Import.to_lower,
                                            'Служебен имейл | Work email  ': Import.to_lower}, ) \
                    .fillna(np.nan).replace([np.nan], [None])
                break
        return report_df

    @staticmethod
    def import_limitations():
        expected_columns = ['COMPANY', 'C_PER_PERSON', 'C_PER_MONTH', 'PREPAID', 'START', 'END', 'DURATION DAYS',
                            'NOTE', 'BGN_PER_HOUR', 'IS_VALID']
        limitations_df = None
        while True:
            tkinter.Tk().withdraw()  # prevents an empty tkinter window from appearing
            file_path = filedialog.askopenfilename(title='Select the file with the Limitations',
                                                   filetypes=[("CSV (comma-separated values) file", ".csv")],
                                                   defaultextension=".csv",
                                                   )
            if not file_path:
                print(f"Please choose the .csv file with the predefined Limitations!")
                continue

            columns_in_selected_file = pd.read_csv(file_path).columns
            column_inconsistency_check = [x for x in zip(expected_columns, columns_in_selected_file) if x[0] != x[1]]
            if column_inconsistency_check:
                print(column_inconsistency_check, sep="\n")
                message = [">>> The following columns are not matching expected order/names: "]
                for items in column_inconsistency_check:
                    expected = items[0]
                    found = items[1]
                    line_to_add = f"The expected column: '{expected}' is not equal to the '{found}' " \
                                  f"column from the chosen file!"
                    message.append(line_to_add)

                print(f"Please check the column's consistency in the chosen file!\n",
                      '\n'.join(message))
                continue
            else:
                limitations_df = pd.read_csv(file_path,
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
                break
        return limitations_df


class Export:
    __export_folder = "exports/"
    __archive_folder = "archives/"

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
    def df_to_csv(dataframe: pd.DataFrame, name: str) -> None:
        """
        Converts the DateFrame with the data to .csv
        :param dataframe: transformed from BaseWebsite._data_dict
        :param name: depends on instance name
        :return: nothing
        """
        dataframe.to_csv(f"{Export.__export_folder}{name}.csv", encoding='utf-8', index=False)

    @staticmethod
    def companies_df_to_excel(name: str, dataframe: pd.DataFrame, column_list: List[str] = None):
        # preparing the new df by company (is_valid == 1) in scope + (is_valid == 0) out of the project scope
        with pd.ExcelWriter(f'{Export.__export_folder}{name}.xlsx', engine='xlsxwriter') as ew:
            # add worksheets for the trainings in scope
            for value in column_list:
                company_filter = (dataframe['company'] == value)

                # use the dataframe filtering for getting data and pass it to the invoice creator
                invoice_data_df_init = dataframe.loc[company_filter &
                                                     (dataframe['is_valid'] == 1)][['nickname', 'type']]

                invoice_data_df = invoice_data_df_init.value_counts()
                rate_per_hour = float(dataframe.loc[company_filter, 'bgn_per_hour'].unique())
                invoice_data_dict = invoice_data_df.to_dict()

                if invoice_data_dict:
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
                new_df.loc[-1, 'total_trainings'] = new_df['training_datetime'].count()
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
            month_stats_emp = mont_df[['nickname', 'company']] \
                .value_counts() \
                .reset_index(name='Trainings per Employee') \
                .sort_values(by='nickname')
            month_stats_emp.loc[-1, 'total'] = month_stats_emp['Trainings per Employee'].sum()
            month_stats_emp.to_excel(ew, sheet_name='by employee',
                                     index=False,
                                     header=['EmployeeID', 'Company', 'trainings', 'Total'],
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
                                            index=True,
                                            index_label=['Company', 'EmployeeID'],
                                            header=['trainings'],
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
                                         index=True,
                                         index_label=['Trainer', 'EmployeeID'],
                                         header=['trainings'],
                                         freeze_panes=(1, 3))
            # add sheet for statistical monthly data for Company's total
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
                                            index=True,
                                            index_label='Company',
                                            header=['BGN', 'Total trainings'],
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
                                                index=True,
                                                index_label='Trainer',
                                                header=['BGN', 'Total trainings'],
                                                freeze_panes=(1, 3))

    @staticmethod
    def stats_full_df_to_excel(name: str, full_df: pd.DataFrame):
        with pd.ExcelWriter(f'{Export.__export_folder}{name}.xlsx', engine='xlsxwriter') as ew:
            # add sheet for statistical annual data by Company with percentage
            # filter_for_actual_contracts_y = (full_df['is_valid'] == 1)
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
                                          header=['Company', 'Total trainings', '%Percentage'],
                                          freeze_panes=(1, 3)
                                          )
            # add sheet for statistical annual data by Company and Year
            annual_stats_by_year = full_df[['company', 'year']]
            annual_stats_by_year = annual_stats_by_year.value_counts(['company', 'year']).unstack()
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
                                             header=['Company', 'Total trainings', '%Percentage'],
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
            # add sheet for statistical annual data by Company and Employees with only 1 training
            annual_stats_by_emp_with_one_training = full_df[
                (full_df['returns_or_not'] == 'only one session')].reset_index()
            annual_stats_by_emp_with_one_training = pd.DataFrame(
                annual_stats_by_emp_with_one_training[['company', 'nickname']].value_counts().to_frame())
            annual_stats_by_emp_with_one_training.loc[-1, 'Grand Total'] = annual_stats_by_emp_with_one_training[
                'count'].sum()
            annual_stats_by_emp_with_one_training.to_excel(ew, sheet_name='by_emp_with_one_training',
                                                           index=True,
                                                           index_label=['Company', 'EmployeeID'],
                                                           header=['Count', 'Grand Total'],
                                                           freeze_panes=(1, 4))

    @staticmethod
    def save_zip_via_browser() -> str:
        tkinter.Tk().withdraw()  # prevents an empty tkinter window from appearing
        target_path = asksaveasfile(filetypes=[("Zip archive", ".zip")],
                                    defaultextension=".zip",
                                    initialfile="all_data.zip",
                                    title='Save all data/reports as Zip file')
        original_file = os.listdir(Export.__archive_folder)[0]
        original_path = fr'{Export.__archive_folder}{original_file}'
        if target_path:
            shutil.copyfile(original_path, target_path.name)
        else:
            print("Please select where to save the reports '.zip' file")
            Export.save_zip_via_browser()
        return target_path


class ZipFiles:
    __folder_path = "archives/"

    @staticmethod
    def zip_export_folder() -> str:
        dt = datetime.now()
        dt_string = dt.strftime("%d-%m-%Y_%H-%M-%S")
        new_file = f"reports_and_invoices_{dt_string}.zip"
        tkinter.Tk().withdraw()  # prevents an empty tkinter window from appearing
        target_dir_path = asksaveasfile(filetypes=[("Zip archive file", ".zip")],
                                        defaultextension=".zip",
                                        initialfile=f"{new_file}",
                                        title='Save all data/reports as Zip file')

        save_as = target_dir_path.name
        with zipfile.ZipFile(f'{save_as}', 'w') as f:
            # exclude manifests files (files starting with _) with glob: [!_]
            for file in glob.glob(f'exports/[!_]*'):
                f.write(file)
        return new_file


class Clearing:

    @staticmethod
    def delete_files_from_folder(extensions: list):
        # Search files with .txt extension in current directory
        for ext in extensions:
            pdf_pattern = f"exports/*.{ext}"
            files = glob.glob(pdf_pattern)
            # deleting the files with listed as argument extensions
            for file in files:
                os.remove(file)

    @staticmethod
    def delete_all_zip_from_archive_folder():
        # Search files with .txt extension in current directory
        pdf_pattern = f"archives/*.zip"
        files = glob.glob(pdf_pattern)
        # deleting the files with zip extension
        for file in files:
            os.remove(file)

# ZipFiles.zip_export_folder()
# ZipFiles.zip_export_invoices()
# ZipFiles.zip_export_reports()
# Clearing.delete_files_from_folder(["csv", "pdf", "xlsx"])
# Clearing.delete_all_zip_from_archive_folder()
# Export.save_zip_via_browser()
# Export.save_zip_via_browser()
