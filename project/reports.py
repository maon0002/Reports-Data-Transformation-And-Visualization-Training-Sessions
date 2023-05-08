# coding: utf8
import gc
from abc import ABC, abstractmethod
from typing import Dict
import pandas as pd
import logging
from project.file_operations import Import, Export, ZipFiles, Clearing
from project.transformations import Transformation
from tqdm import tqdm

logging.basicConfig(filename='info.log', encoding='utf-8',
                    level=logging.INFO,
                    format=u'%(asctime)s %(levelname)-8s %(message)s',
                    datefmt='%d-%m-%Y %H:%M:%S',
                    filemode="w",)


class BaseReport(ABC):
    """An abstract base class used to represent a report

    Attributes
    ----------
    name : str
        a string with the report name
    path : str
        a string with the target export path
    export_function : str
        the function from project.transformations Export which handle the report exporting

    Methods
    -------
    build_report_base() -> Dict[str, pd.DataFrame]:
        Imports two standard files from the local PC which are
        needed fundament/base for the further data validation/transformation
        and returns dataframes objects for different reporting purposes

    export_report(self):
        An abstract method which links particular Export function to a
        child class instance and trigger the final formatting/conversion
        and the exporting to a file
    """
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
    def build_report_base() -> Dict[str, pd.DataFrame]:
        """
        Transform the monthly and the annual df columns
        """
        original_report: pd.DataFrame = Import.import_report()
        limitations_file: pd.DataFrame = Import.import_limitations()
        dataframes_dict = Transformation.main(original_report, limitations_file)
        return dataframes_dict

    def get_function_by_name(self):
        function = [f[1] for f in BaseReport._export_functions_dict.items()
                    if f[0] == self.export_function][0]
        return function

    @abstractmethod
    def export_report(self):
        ...

    def __repr__(self):
        return f"{self.name}"

    def __str__(self):
        return f"{self.name.lower()}"


class Report(BaseReport):
    """The class used to represent a single reports based on a
    particular dataframe

    Attributes
    ----------
    name : str
        a string with the report name
    path : str
        a string with the target export path
    export_function : str
        the function from project.transformations Export which handle the report exporting
    df_name : str
        the dataframe variable name which links the dataframe object itself

    Methods
    -------
    export_report(self):
        It search for the right function from a dictionary
        and execute it with the needed parameters

    def get_dataframe_obj(self):
        Search for the right dataframe using the object attribute df_name
        in the dataframes dictionary
    """

    def __init__(self, name, path: str, export_function: str, df_name: str = ""):
        super().__init__(name, path, export_function)
        self.df_name = df_name

    def get_dataframe_obj(self):
        dataframe_obj = [d[1] for d in dataframes_dictionary.items()
                         if d[0] == self.df_name][0]

        return dataframe_obj

    def export_report(self):
        """
        Find the needed dataframe and execute the right function for this class object
        """
        logging.info(f"Initiate exporting of the {self.name} report")
        function = self.get_function_by_name()
        df_obj = self.get_dataframe_obj()
        function(self.name, df_obj, self.path)
        logging.info(f"Exporting of the {self.name} report was successfully done")


class BulkReport(BaseReport):
    """The class used to represent the reports which are
    filtered by given value for the data in the monthly new formatted dataframe.
    Reports that are extracted by company or calendar values in a separate reports bu in a bulk

   Attributes
   ----------
   name : str
       a string with the report name
   path : str
       a string with the target export path
   export_function : str
       the function from project.transformations Export which handle the report exporting

   Methods
   -------
   export_report(self):
       It search for the right function from a dictionary
        and execute it with the needed parameters

   def get_dataframe_obj(self):
       Search for the right dataframe using the object attribute df_name
       in the dataframes dictionary
   """

    def __init__(self, name, path: str, export_function: str):
        super().__init__(name, path, export_function)

    def export_report(self):
        """
        Find and execute the right function for this class object
        """
        logging.info(f"Initiate exporting of the {self.name} report")
        function = self.get_function_by_name()
        function(self.path)
        logging.info(f"Exporting of the {self.name} report was successfully done")


class MultiReport(BaseReport):
    """Class used to represent a multi-reports (All-in-one)

        Attributes
        ----------
        name : str
            a string with the report name
        path : str
            a string with the target export path
        export_function : str
            the function from project.transformations Export which handle the report exporting

        Methods
        -------
        export_report(self):
            It search for the right function from a dictionary
            and execute it with the needed parameters
        """

    def __init__(self, name, path: str, export_function: str, df_dict: Dict[str, pd.DataFrame]):
        super().__init__(name, path, export_function)
        self.df_dict = df_dict

    def export_report(self):
        """
        Find and execute the right function for this class object
        """
        logging.info(f"Initiate exporting of the {self.name} report")
        function = self.get_function_by_name()
        function(self.name, self.df_dict, self.path)
        logging.info(f"Exporting of the {self.name} report was successfully done")


# Remove all files in the 'exports' folder and sub-folders saved during previous runs
Clearing.delete_files_from_export_subfolders()

# Get a dictionary with the based on the imported initial report dataframes
dataframes_dictionary = BaseReport.build_report_base()

# Create all report instances
raw_full = Report("Raw_Full", "exports/for_reference/", "df_to_csv", "full_raw_report_df")
raw_mont = Report("Raw_Mont", "exports/for_reference/", "df_to_csv", "monthly_raw_report_df")
new_full = Report("New_Full", "exports/for_reference/", "df_to_csv", "new_full_data_df")
new_mont = Report("New_Mont", "exports/for_reference/", "df_to_csv", "new_monthly_data_df")
companies = Report("Companies", "exports/", "companies_df_to_excel", "new_monthly_data_df")
companies_out_of_scope = Report("Companies_Out_Of_Scope", "exports/", "companies_df_to_excel", "new_monthly_data_df")
trainers = Report("Trainers", "exports/", "trainers_df_to_excel", "new_monthly_data_df")
generic = MultiReport("Generic", "exports/for_reference/", "generic_df_to_excel", dataframes_dictionary)
stats_mont = Report("Stats_Mont", "exports/for_reference/", "stats_mont_df_to_excel", "new_monthly_data_df")
stats_full = Report("Stats_Full", "exports/for_reference/", "stats_full_df_to_excel", "new_full_data_df")
by_calendar = BulkReport("by_calendar", "exports/from_templates/by_calendar", "convert_docx_to_pdf")
by_company = BulkReport("by_company", "exports/from_templates/by_company", "convert_docx_to_pdf")
by_company_x = BulkReport("by_company_x", "exports/from_templates/by_company_x", "convert_docx_to_pdf")

# Get a list of all report instances and initiate progress tracking
report_instances = [r for r in gc.get_objects() if isinstance(r, BaseReport)]
for i in tqdm(range(len(report_instances))):
    report_instances[i].export_report()

# Zip all files and folders in 'exports' folder
ZipFiles.zip_export_folder()
