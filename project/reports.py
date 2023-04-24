# coding: utf8
from abc import ABC, abstractmethod
from typing import List
import pandas as pd
import logging
from project.dataframes import BaseDataframe
from project.file_operations import Import, Export
from project.collections import Collection
from project.transformations import Transformation

logging.basicConfig(filename='info.log', encoding='utf-8',
                    level=logging.INFO,
                    format=u'%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%d-%m-%Y %H:%M:%S'
                    )


class BaseReport(ABC):

    def __init__(self, name):
        self.name = name

    @abstractmethod
    def build_report(self, dataframe: pd.DataFrame, limitations_df: pd.DataFrame):
        ...


class Report(BaseReport):
    """

    """

    def __init__(self, name: str):
        super().__init__(name)
        self.name = name
        self.main_column = None
        self.needed_columns = None
        self.dataframe: pd.DataFrame = None
        self.dataframes_list: List[BaseDataframe] = []
        self.new_headers = None

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

    def build_report(self, dataframe: pd.DataFrame, limitations_df: pd.DataFrame):
        """
        Transform the monthly and the annual df columns
        """
        raw_full_df, raw_mont_df, new_full_df, new_mont_df, total_trainers_df = Transformation.main(dataframe,
                                                                                                    limitations_df)
        raw_full_report = Report("Raw_Full")
        raw_full_report.dataframe = raw_full_df
        Export.df_to_csv(raw_full_report.dataframe, raw_full_report.name)
        raw_mont_report = Report("Raw_Mont")
        raw_mont_report.dataframe = raw_mont_df

        companies_report = Report("Companies", )
        companies_out_of_scope_report = Report("Companies_Out_Of_Scope", )
        new_full_report = Report("New_Full", )
        new_mont_report = Report("New_Mont", )
        stats_full_report = Report("Stats_Full", )
        stats_mont_report = Report("Stats_Mont", )
        trainers_report = Report("Trainers", )
        return []

    def __repr__(self):
        return f"{self.name}"


original_report = Import.import_report('imports/schedule2023-04-18.csv')
limitations_file = Import.import_limitations('imports/limitations.csv')

Report.build_report(original_report, limitations_file)
