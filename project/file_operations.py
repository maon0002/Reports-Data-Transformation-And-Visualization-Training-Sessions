# coding: utf8
import logging
from datetime import datetime
import zipfile
import os
from typing import List, Dict
import numpy as np
import pandas as pd
from project.invoices import BaseInvoice
import tkinter
from tkinter import filedialog
from tkinter.filedialog import asksaveasfile
from project._collections import Collection
from project.templates import ReportFromTemplate
import subprocess


class Import:
    """Class used to provide functions related to the file importing

        Attributes
        ----------
        No attributes

        Methods
        -------
        to_lower(x):
            Used as a function for the email series/column to lowercase tha values during the import.
            * 'Email' and 'Служебен имейл | Work email  '

        import_report() -> pd.DataFrame:
            Import the general/initial report from the platform for reservations/booking
            and create a dataframe with the full data for further transformations and validation

        import_limitations() -> pd.DataFrame:
            Import the limitation report that contains company contracts data, payment rate per hour etc.
            and create a dataframe

    """

    @staticmethod
    def to_lower(x):
        """
        Used as a function for the email series/column to lowercase tha values during the import.
        * 'Email' and 'Служебен имейл | Work email  '

        :param x:
        :return:
        """
        return x.lower() if isinstance(x, str) else x

    @staticmethod
    def import_report() -> pd.DataFrame:
        """
        Import the general/initial report from the platform for reservations/booking
        and create a dataframe with the full data for further transformations and validation

        :return: dataframe 'report_df'
        """

        # uses file explorer browser to find and select the
        # general/initial report and validate the chose
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
            logging.info(f"The user select the following path: '{file_path}' for initial report importing")
            # check for the consistency of the columns comparing them with a list from _collections
            columns_in_selected_file = pd.read_csv(file_path).columns
            column_inconsistency_check = [x for x in zip(Collection.report_expected_columns(), columns_in_selected_file)
                                          if x[0] != x[1]]

            # print and log the inconsistencies
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
                logging.info(f"***Column inconsistencies for the reservations report:\n", '\n'.join(message))
                continue

            # if choice is valid and there are no inconsistencies
            # import and transform the report data to a dataframe
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
                logging.info(f"General/initial report file imported successfully and "
                             f"limitations_df dataframe was created")
                break
        return report_df

    @staticmethod
    def import_limitations() -> pd.DataFrame:
        """
        Import the limitation report that contains company contracts data, payment rate per hour etc.
        and create a dataframe
        :return: dataframe 'limitations_df'
        """
        # uses file explorer browser to find and select the
        # limitations.csv and validate the chose
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

            logging.info(f"The user select the following path: '{file_path}' for limitations importing")

            # check for the consistency of the columns comparing them with a list from _collections
            columns_in_selected_file = pd.read_csv(file_path).columns
            column_inconsistency_check = [x for x in zip(Collection.limitations_expected_columns(),
                                                         columns_in_selected_file) if x[0] != x[1]]

            # print and log the inconsistencies
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
                logging.info(f"***Column inconsistencies for the limitations report:\n", '\n'.join(message))
                continue

            # if choice is valid and there are no inconsistencies
            # import and transform the report data to a dataframe
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
                logging.info(f"Limitations file imported successfully and limitations_df dataframe was created")
                break
        return limitations_df


class Export:
    """Class used to provide functions related to the file importing

        Attributes
        ----------
        No attributes

        Methods
        -------
        df_to_csv(name: str, dataframe: pd.DataFrame, path: str) -> None:
            Converts the DateFrame with the data to .csv

        companies_df_to_excel(name: str, dataframe: pd.DataFrame, path: str) -> None:
            Takes data from the new monthly dataframe, filters by each company
            and export them in separate sheets, also creates an invoice and an additional reports
            based on a pre-formatted .docx file

        trainers_df_to_excel(name: str, dataframe: pd.DataFrame, path: str) -> None:
            Takes data from the new monthly dataframe, filters by each trainer
            and export them in separate sheets, also creates additional reports
            based on a pre-formatted .docx file

        generic_df_to_excel(name: str, dictionary: Dict[str, pd.DataFrame], path: str) -> None:
            Creates one file with multiple dataframes in separate sheets in case of data review
            Serves class MultiReport

        stats_mont_df_to_excel(name: str, dataframe: pd.DataFrame, path: str) -> None:
            Uses the monthly dataframe to aggregate/pivot the data by:
             - employee
             - company
             - trainer
             - company total
             - trainer total
             and save every dataframe as a separate sheet

        stats_full_df_to_excel(name: str, dataframe: pd.DataFrame, path: str) -> None:
            Uses the full/annual dataframe to aggregate/pivot the data by:
             - employee
             - company
             - trainer
             - company total
             - trainer total
             and save every dataframe as a separate sheet

        convert_docx_to_pdf(files_path) -> None:
            Uses LibreOffice to convert the generated .docx reports to PDF
            Serves to class BulkReport

    """

    @staticmethod
    def df_to_csv(name: str, dataframe: pd.DataFrame, path: str) -> None:
        """
        Converts the DateFrame with the data to .csv
        :param path:
        :param dataframe: transformed from BaseWebsite._data_dict
        :param name: depends on instance name
        :return: nothing
        """
        dataframe.to_csv(f"{path}{name}.csv", encoding='utf-8', index=False)

    @staticmethod
    def companies_df_to_excel(name: str, dataframe: pd.DataFrame, path: str) -> None:
        """
        Takes data from the new monthly dataframe, filters by each company
        and export them in separate sheets, also creates an invoice and an additional reports
        based on a pre-formatted .docx file
        Serves class Report.

        :param name: 'Companies' or 'Companies_out_of_scope'
        :param dataframe: new_monthly_data_df
        :param path: relative path where the file must be saved
        :return: None
        """
        # preparing the new df by company (is_valid == 1) in scope + (is_valid == 0) out of the project scope
        if name == "Companies":
            column_list = Collection.company_report_list(dataframe)
        else:
            column_list = Collection.company_report_list_other(dataframe)
        with pd.ExcelWriter(f'{path}{name}.xlsx', engine='xlsxwriter') as ew:
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
                    # make an invoice for the company
                    BaseInvoice.create_invoice(value, rate_per_hour, invoice_data_dict)

                    # add variables to give the needed inf for creation of the .docx templates
                    new_df = dataframe.loc[company_filter, 'nickname'].value_counts().reset_index(name='count')
                    start_date = dataframe.loc[company_filter &
                                               (dataframe['is_valid'] == 1)][['training_datetime']].iloc[0].astype(str)
                    start_date = str(start_date[0][0:11])
                    end_date = dataframe.loc[company_filter &
                                             (dataframe['is_valid'] == 1)][['training_datetime']].iloc[-1].astype(str)
                    end_date = str(end_date[0][0:11])
                    total_hours = new_df[['count']].sum().iloc[-1]

                    # create the .docx templates with the company data
                    ReportFromTemplate.create_by_company_report_from_docx_template(
                        value, new_df, total_hours, start_date, end_date)
                    ReportFromTemplate.create_by_company_x_report_from_docx_template(
                        value, new_df, total_hours, start_date, end_date)

                # aggregate the company data, add 'count' and 'total' columns
                new_df = dataframe.loc[company_filter, 'nickname'].value_counts().reset_index(name='count')
                new_df.loc[-1, 'total'] = new_df['count'].sum()

                # export/ save the dataframe to excel as sheet with company name as a sheet name
                new_df.to_excel(ew, sheet_name=value, header=['Employee ID', 'Bookings', 'Total'],
                                index=False)

    @staticmethod
    def trainers_df_to_excel(name: str, dataframe: pd.DataFrame, path: str) -> None:
        """
        Takes data from the new monthly dataframe, filters by each trainer
        and export them in separate sheets, also creates additional reports
        based on a pre-formatted .docx file.
        Serves class Report.

        :param name: the name for the .excel file
        :param dataframe: new_monthly_data_df
        :param path: relative path where the file must be saved
        :return: None
        """

        # take the unique values in the 'trainer' pd.Series (trainers names*)
        column_list = Collection.trainers_report_list(dataframe)

        # use the manager for the .xlsx file building and saving
        with pd.ExcelWriter(f'{path}{name}.xlsx', engine='xlsxwriter') as ew:
            for value in column_list:
                new_df = dataframe[(dataframe['trainer'] == value)][
                    ['company', 'training_datetime', 'employee_names', 'short_type', 'bgn_per_hour']]

                # add two columns for totals
                total_hours = new_df['training_datetime'].count()
                total_pay = new_df['bgn_per_hour'].sum()

                # make additional report via .docx template
                ReportFromTemplate.create_by_calendar_report_from_docx_template(value, new_df, total_hours, total_pay)

                new_df.loc[-1, 'total_trainings'] = total_hours
                new_df.loc[-1, 'total_pay'] = f"{total_pay:.2f}" + ".лв"

                # export/ save the dataframe to excel as sheet with trainer name as a sheet name
                new_df.to_excel(ew, sheet_name=value,
                                header=['Компания', 'Време на тренинга', 'Имена на служител', 'Вид', 'Лв на час',
                                        'Общо тренинги', 'Общо възнаграждение'], index=False)

    @staticmethod
    def generic_df_to_excel(name: str, dictionary: Dict[str, pd.DataFrame], path: str) -> None:
        """
        Creates one file with multiple dataframes in separate sheets in case of data review
        Serves class MultiReport

        :param name: the name for the .excel file
        :param dictionary: dictionary with the dataframes which will be saved as separate excel sheets
        :param path: relative path where the file must be saved
        :return: None
        """

        # use the manager for the .xlsx file building and saving
        with pd.ExcelWriter(f'{path}{name}.xlsx', engine='xlsxwriter') as ew:
            for item in dictionary.items():
                df = item[1]
                df_name = item[0]
                if df_name in Collection.generic_report_list():
                    df.to_excel(ew, sheet_name=str(df_name), index=False)
            month_describe = pd.DataFrame(dictionary["new_monthly_data_df"].describe())
            annual_describe = pd.DataFrame(dictionary["new_full_data_df"].describe())
            month_describe.to_excel(ew, sheet_name='month_describe')
            annual_describe.to_excel(ew, sheet_name='annual_describe')

    @staticmethod
    def stats_mont_df_to_excel(name: str, dataframe: pd.DataFrame, path: str) -> None:
        """
        Uses the monthly dataframe to aggregate/pivot the data by:
         - employee
         - company
         - trainer
         - company total
         - trainer total
         and save every dataframe as a separate sheet

        :param name: the name for the .excel file
        :param dataframe: new_monthly_data_df
        :param path: relative path where the file must be saved
        :return: None
        """

        # use the manager for the .xlsx file building and saving
        with pd.ExcelWriter(f'{path}{name}.xlsx', engine='xlsxwriter') as ew:
            # add sheet for statistical monthly data by Employee
            month_stats_emp = dataframe[['nickname', 'company']] \
                .value_counts() \
                .reset_index(name='Trainings per Employee') \
                .sort_values(by='nickname')
            month_stats_emp.loc[-1, 'total'] = month_stats_emp['Trainings per Employee'].sum()
            month_stats_emp.to_excel(ew, sheet_name='by employee',
                                     index=False,
                                     header=['EmployeeID', 'Company', 'trainings', 'Total'],
                                     freeze_panes=(1, 4))

            # add sheet for statistical monthly data by Company
            month_stats_comp = dataframe[['company', 'nickname', 'concat_emp_company']]
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
            month_stats_trainer = dataframe[['trainer', 'nickname', 'concat_emp_company']]
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
            month_stats_comp_total = dataframe[['company', 'concat_emp_company', 'bgn_per_hour']]
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
            month_stats_trainer__total = dataframe[['trainer', 'bgn_per_hour', 'is_valid']]
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
    def stats_full_df_to_excel(name: str, dataframe: pd.DataFrame, path: str) -> None:
        """
        Uses the full/annual dataframe to aggregate/pivot the data by:
         - employee
         - company
         - trainer
         - company total
         - trainer total
         and save every dataframe as a separate sheet

        :param name: the name for the .excel file
        :param dataframe: new_monthly_data_df
        :param path: relative path where the file must be saved
        :return:
        """
        with pd.ExcelWriter(f'{path}{name}.xlsx', engine='xlsxwriter') as ew:
            # add sheet for statistical annual data by Company with percentage
            # filter_for_actual_contracts_y = (full_df['is_valid'] == 1)
            annual_stats_general1 = dataframe[['company']].value_counts().reset_index()
            annual_stats_general2 = (dataframe[['company']].
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
                                          freeze_panes=(1, 3))

            # add sheet for statistical annual data by Company and Year
            annual_stats_by_year = dataframe[['company', 'year']]
            annual_stats_by_year = annual_stats_by_year.value_counts(['company', 'year']).unstack()
            annual_stats_by_year['Total'] = annual_stats_by_year.agg("sum", axis='columns')
            annual_stats_by_year.loc[-1, 'Grand Total'] = annual_stats_by_year['Total'].sum()
            annual_stats_by_year.to_excel(ew, sheet_name='by_year',
                                          index=['company'],
                                          index_label=['Company'],
                                          freeze_panes=(1, 6))

            # add sheet for statistical annual data by Unique EmployeeIDs with percentage
            annual_stats_unique_emp0 = dataframe[['company', 'nickname']].value_counts().to_frame().reset_index()
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
            annual_stats_by_emp_by_year = dataframe[['company', 'nickname', 'year']]

            annual_stats_by_emp_by_year = annual_stats_by_emp_by_year.value_counts(['company', 'nickname', 'year']) \
                .unstack()

            annual_stats_by_emp_by_year['Total'] = annual_stats_by_emp_by_year.agg("sum", axis='columns')

            annual_stats_by_emp_by_year.loc[-1, 'Grand Total'] = annual_stats_by_year['Total'].sum()

            annual_stats_by_emp_by_year.to_excel(ew, sheet_name='by_emp_by_year',
                                                 index=['company', 'nickname'],
                                                 index_label=['Company', 'EmployeeID'],
                                                 freeze_panes=(1, 7))

            # add sheet for statistical annual data by Company and Employees with only 1 training
            annual_stats_by_emp_with_one_training = \
                dataframe[(dataframe['returns_or_not'] == 'only one session')].reset_index()

            annual_stats_by_emp_with_one_training = pd.DataFrame(
                annual_stats_by_emp_with_one_training[['company', 'nickname']].value_counts().to_frame())

            annual_stats_by_emp_with_one_training.columns = ['count']

            annual_stats_by_emp_with_one_training.loc[-1, 'Grand Total'] = \
                annual_stats_by_emp_with_one_training['count'].sum()

            annual_stats_by_emp_with_one_training.to_excel(ew, sheet_name='by_emp_with_one_training',
                                                           index=True,
                                                           index_label=['Company', 'EmployeeID'],
                                                           header=['Count', 'Grand Total'],
                                                           freeze_panes=(1, 4))

    @staticmethod
    def convert_docx_to_pdf(files_path) -> None:
        """
        Uses LibreOffice to convert the generated .docx reports to PDF
        Serves to class BulkReport

        :param files_path: relative path where the file must be saved
        :return: None
        """

        # path to the engine
        path_to_office = r"C:\Program Files\LibreOffice\program\soffice.exe"
        # verify the path using getcwd()
        cwd = os.getcwd()
        files = os.listdir(files_path)
        for file in files:
            if os.path.isfile(files_path + "/" + file):
                # path with files to convert
                source_folder = fr"{files_path}"

                # path with pdf files
                output_folder = fr"PDFs"

                # changing directory to source
                os.chdir(source_folder)

                # assign and running the command of converting files through LibreOffice
                command = f"\"{path_to_office}\" --convert-to pdf  --outdir \"{output_folder}\" *.*"
                subprocess.run(command)
        os.chdir(cwd)


class ZipFiles:
    """
    Class used to provide functions related to the archive files

        Attributes
        ----------
        No attributes

        Methods
        -------
        zip_export_folder() -> str:
            Collects filenames/file paths and add the files into a .zip archive file.

    """

    @staticmethod
    def zip_export_folder() -> None:
        """
        Collects filenames/file paths and add the files into a .zip archive file.

        :return: None
        """

        dt = datetime.now()
        dt_string = dt.strftime("%d-%m-%Y_%H-%M-%S")
        new_file = f"reports_and_invoices_{dt_string}.zip"

        # use file browser for the path selection and validate choice/actions
        while True:
            tkinter.Tk().withdraw()  # prevents an empty tkinter window from appearing
            target_dir_path = asksaveasfile(filetypes=[("Zip archive file", ".zip")],
                                            defaultextension=".zip",
                                            initialfile=f"{new_file}",
                                            title='Save all data/reports as Zip file')

            if not target_dir_path:
                print("Please select folder for the output .zip file!")
                continue

            logging.info(f"The user select the following path for saving the .zip: '{target_dir_path}'")
            save_as = target_dir_path.name

            # get list of all files in 'exports' folder
            with zipfile.ZipFile(f'{save_as}', 'w') as f:
                for root, dirs, files in os.walk("exports"):
                    for file in files:
                        f.write(os.path.join(root, file))
                    for directory in dirs:
                        f.write(os.path.join(root, directory))
            logging.info(f"The .zip files with all of the reports was created and saved")
            break


class Clearing:
    """
    Class used to provide functions related to the file removing

        Attributes
        ----------
        No attributes

        Methods
        -------
        list_files(folder_relative_path) -> List[str]:
            List files and folders/sub folders in given relative path

        delete_files_from_export_subfolders():
            Receives list with file paths and names and removes all non-directories
    """

    @staticmethod
    def list_files(folder_relative_path) -> List[str]:
        """
        List files and folders/sub folders in given relative path

        :param folder_relative_path:
        :return: list with the file paths and names
        """
        logging.info(f"Removing of the files from '{folder_relative_path}' was initiated!")
        items_list = []
        for root, dirs, files in os.walk(folder_relative_path, topdown=False):
            for name in files:
                items_list.append(os.path.join(root, name))
            for name in dirs:
                items_list.append(os.path.join(root, name))
        logging.info(f"Removing of the files from '{folder_relative_path}' completed!")
        return items_list

    @staticmethod
    def delete_files_from_export_subfolders():
        """
        Receives list with file paths and names and removes all non-directories
        :return:
        """
        clearing_list = Clearing.list_files("exports")

        for item in clearing_list:
            if os.path.isfile(item):
                os.remove(item)
