# coding: utf8
from abc import ABC, abstractmethod
import pandas as pd
import logging
from project.file_operations import Import, Export, ZipFiles, Clearing
from project._collections import Collection
from project.transformations import Transformation

logging.basicConfig(filename='info.log', encoding='utf-8',
                    level=logging.INFO,
                    format=u'%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%d-%m-%Y %H:%M:%S')


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

    @staticmethod
    def build_report(dataframe: pd.DataFrame, limitations_df: pd.DataFrame):
        """
        Transform the monthly and the annual df columns
        """
        dataframes_dict = Transformation.main(dataframe, limitations_df)
        total_trainings_df, report_trainers_df, new_monthly_data_df, new_full_data_df, limitations_df, flags_data_df, \
        full_raw_report_df, monthly_raw_report_df = list(dataframes_dict.values())

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


original_report = Import.import_report()
limitations_file = Import.import_limitations()

Report.build_report(original_report, limitations_file)
ZipFiles.zip_export_folder()

Clearing.delete_files_from_folder(["csv", "pdf", "xlsx"])
Clearing.delete_all_zip_from_archive_folder()
