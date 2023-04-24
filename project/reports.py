# coding: utf8
from abc import ABC, abstractmethod
from typing import List
import pandas as pd
import logging
from project.dataframes import BaseDataframe
from project.file_operations import Import, Export
from project._collections import Collection
from project.transformations import Transformation

logging.basicConfig(filename='info.log', encoding='utf-8',
                    level=logging.INFO,
                    format=u'%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%d-%m-%Y %H:%M:%S'
                    )


class BaseReport(ABC):

    def __init__(self, name):
        self.name = name

    @staticmethod
    @abstractmethod
    def build_report(dataframe: pd.DataFrame, limitations_df: pd.DataFrame):
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

    @staticmethod
    def build_report(dataframe: pd.DataFrame, limitations_df: pd.DataFrame):
        """
        Transform the monthly and the annual df columns
        """
        dataframes_dict = Transformation.main(dataframe, limitations_df)
        total_trainings_df, report_trainers_df, new_monthly_data_df, new_full_data_df, limitations_df, flags_data_df, \
        full_raw_report_df, monthly_raw_report_df = list(dataframes_dict.values())

        # raw_full_df, raw_mont_df, new_full_df, new_mont_df, total_trainings_df, report_trainers_df, flags_df = \
        #     Transformation.main(dataframe, limitations_df)

        Export.df_to_csv(full_raw_report_df, "Raw_Full")
        Export.df_to_csv(monthly_raw_report_df, "Raw_Mont")
        Export.df_to_csv(new_full_data_df, "New_Full")
        Export.df_to_csv(new_monthly_data_df, "New_Mont")
        Export.companies_df_to_excel("Companies", new_monthly_data_df,
                                     Collection.company_report_list(new_monthly_data_df))
        Export.companies_df_to_excel("Companies_Out_Of_Scope", new_monthly_data_df,
                                     Collection.company_report_list_other(new_monthly_data_df))
        Export.trainers_df_to_excel("Trainers", new_monthly_data_df,
                                    Collection.trainers_report_list(new_monthly_data_df))
        Export.generic_df_to_excel("Generic", dataframes_dict)

        Export.stats_mont_df_to_excel("Stats_Mont", new_monthly_data_df)
        Export.stats_full_df_to_excel("Stats_Full", new_full_data_df)

        return []

    def __repr__(self):
        return f"{self.name}"


original_report = Import.import_report('imports/schedule2023-04-18.csv')
limitations_file = Import.import_limitations('imports/limitations.csv')

Report.build_report(original_report, limitations_file)
