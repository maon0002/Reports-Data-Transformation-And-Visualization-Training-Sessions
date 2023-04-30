# coding: utf8
import inspect
from abc import ABC, abstractmethod
import pandas as pd
import logging
from project.file_operations import Import, Export, ZipFiles, Clearing
from project.transformations import Transformation

logging.basicConfig(filename='info.log', encoding='utf-8',
                    level=logging.INFO,
                    format=u'%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%d-%m-%Y %H:%M:%S')


class BaseReport(ABC):
    _export_functions_dict = {
        "df_to_csv": Export.df_to_csv,
        "companies_df_to_excel": Export.companies_df_to_excel,
        "trainers_df_to_excel": Export.trainers_df_to_excel,
        "generic_df_to_excel": Export.generic_df_to_excel,
        "stats_mont_df_to_excel": Export.stats_mont_df_to_excel,
        "stats_full_df_to_excel": Export.stats_full_df_to_excel,
        "convert_docx_to_pdf": Export.convert_docx_to_pdf
    }

    def __init__(self, name, path: str, export_function: str):
        self.name = name
        self.path = path
        self.export_function = export_function

    @staticmethod
    def build_report_base(dataframe: pd.DataFrame, limitations_df: pd.DataFrame):
        """
        Transform the monthly and the annual df columns
        """
        df_dict = Transformation.main(dataframe, limitations_df)
        return df_dict

    # @staticmethod
    @abstractmethod
    def export_report(self):
        ...


class Report(BaseReport):
    """

    """

    def __init__(self, name, file_path: str, export_function: str, df_name: str = ""):
        super().__init__(name, file_path, export_function)
        self.df_name = df_name

    def get_dataframe_obj(self):
        dataframe_obj = [d[1] for d in dataframes_dict.items() if d[0] == self.df_name][0]
        return dataframe_obj

    def get_function_by_name(self):
        function = [f[1] for f in BaseReport._export_functions_dict.items() if f[0] == self.export_function][0]
        return function

    # @staticmethod
    def export_report(self):
        """
        Transform the monthly and the annual df columns
        """
        function = self.get_function_by_name()
        function_arguments = str(inspect.signature(function))

        if "files_path" in function_arguments:
            function(self.path)
        elif "dictionary" in function_arguments:
            function(self.name, dataframes_dict, self.path)
        else:
            df_obj = self.get_dataframe_obj()
            function(self.name, df_obj, self.path)

    def __repr__(self):
        return f"{self.name}"


Clearing.delete_files_from_export_subfolders()

original_report = Import.import_report()
limitations_file = Import.import_limitations()

dataframes_dict = BaseReport.build_report_base(original_report, limitations_file)

raw_full = Report("Raw_Full", "exports/for_reference/", "df_to_csv", "full_raw_report_df")
raw_mont = Report("Raw_Mont", "exports/for_reference/", "df_to_csv", "monthly_raw_report_df")
new_full = Report("New_Full", "exports/for_reference/", "df_to_csv", "new_full_data_df")
new_mont = Report("New_Mont", "exports/for_reference/", "df_to_csv", "new_monthly_data_df")
companies = Report("Companies", "exports/", "companies_df_to_excel", "new_monthly_data_df")
companies_out_of_scope = Report("Companies_Out_Of_Scope", "exports/", "companies_df_to_excel", "new_monthly_data_df")
trainers = Report("Trainers", "exports/", "trainers_df_to_excel", "new_monthly_data_df")

generic = Report("Generic", "exports/for_reference/", "generic_df_to_excel")

stats_mont = Report("Stats_Mont", "exports/for_reference/", "stats_mont_df_to_excel", "new_monthly_data_df")
stats_full = Report("Stats_Full", "exports/for_reference/", "stats_full_df_to_excel", "new_full_data_df")

raw_full.export_report()
raw_mont.export_report()
new_full.export_report()
new_mont.export_report()
companies.export_report()
companies_out_of_scope.export_report()
trainers.export_report()
generic.export_report()
stats_mont.export_report()
stats_full.export_report()

by_calendar = Report("by_calendar", "exports/from_templates/by_calendar", "convert_docx_to_pdf")
by_company = Report("by_company", "exports/from_templates/by_company", "convert_docx_to_pdf")
by_company_x = Report("by_company_x", "exports/from_templates/by_company_x", "convert_docx_to_pdf")

by_calendar.export_report()
by_company.export_report()
by_company_x.export_report()

ZipFiles.zip_export_folder()


