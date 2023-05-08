import re
from datetime import datetime
import pandas as pd
from project._collections import Collection
from project.dataframes import BaseDataframe


class Transformation:
    """Class used to provide functions for the data transformation control

        Attributes
        ----------
        No attributes

        Methods
        -------
        annual_to_monthly_report_df(report_dataframe: pd.DataFrame, datetime_format: str) -> pd.DataFrame:
            The function takes a piece from the data
            which is related only for the chosen from the user month
            and creates the raw dataframe for the monthly reports: 'monthly_raw_report_df'

        main(dataframe: pd.DataFrame, limitations_dataframe: pd.DataFrame) -> dict:
            Uses the original .csv imported reports (initial report and the limitations files)
            to control the transformation of the dataframes
            (BaseDataframe functions + adding additional series/columns)
        """

    @staticmethod
    def annual_to_monthly_report_df(report_dataframe: pd.DataFrame,
                                    datetime_format: str) -> pd.DataFrame:
        """
        The function takes a piece from the data
        which is related only for the chosen from the user month
        and creates the main dataframe for the monthly reports

        :param report_dataframe:
        :param datetime_format:
        :return: dataframe with records/rows/lines only for the chosen month
        """

        correct_input = False
        chosen_month = None
        # Make variables for the first and the last period(month-year)
        # to use them as an example range
        min_datetime = report_dataframe['start_time'].iloc[0]
        min_date_str = str(min_datetime)[:7]
        max_datetime = report_dataframe['start_time'].iloc[-1]
        max_date_str = str(max_datetime)[:7]

        # Take an user input that defines the period in scope and validate it
        while True:
            if correct_input:
                break
            chosen_month = input(f"Please chose the Year and the Month (YYYY-MM) "
                                 f"between '{min_date_str}' and '{max_date_str}' \n"
                                 f"for the monthly reports in a correct format\n"
                                 f"*'2023-02' for example is February 2023:   ")

            correct_input = re.search(r'(\d{4}-\d{2})', chosen_month.strip())

        # transform the period to an object for further filtering use
        chosen_month_obj = datetime.strptime(chosen_month.strip() +
                                             "-01 00:00:00",
                                             datetime_format)

        chosen_month = pd.Series(chosen_month_obj).dt.to_period('M')

        # filter the data fo only one particular period (month from a year)
        monthly_data_df = report_dataframe.loc[report_dataframe['start_time'].
            dt.to_period('M').
            isin(chosen_month)]. \
            reset_index(drop=True)
        return monthly_data_df

    @staticmethod
    def main(dataframe: pd.DataFrame,
             limitations_dataframe: pd.DataFrame
             ) -> dict:
        """
        Uses the original .csv imported reports
        (initial report and the limitations files)
        to control the transformation of the dataframes
        (BaseDataframe functions + adding additional series/columns)

        :param dataframe:
        :param limitations_dataframe:
        :return: dictionary with dataframes for further reporting use
        """
        flags_data_df = pd.DataFrame(Collection.flags_dict())
        limitations_df = BaseDataframe.limitations_func(limitations_dataframe)
        df = BaseDataframe.rename_original_report_columns(dataframe)

        # add column for issues
        df['flags'] = ""
        df['emp_names_input'] = '|' + df['f_name'] + '|' + df['l_name'] + '|'

        # rough cleaning of the columns data
        df = BaseDataframe.unwanted_chars(df)
        df = BaseDataframe.trim_all_columns(df)

        # transliteration of the employee names from cyrillic to latin chars
        df["first_name"] = BaseDataframe.transliterate_bg_to_en(df, "f_name", "first_name")
        df["last_name"] = BaseDataframe.transliterate_bg_to_en(df, "l_name", "last_name")

        # using first, last name and email parts
        df = BaseDataframe.nickname(df)

        # get only the company name and if the training was IN PERSON/LIVE or ONLINE
        df = BaseDataframe.company_subtraction(df)

        # merge the columns from limitations_df to the monthly/annual df
        df = pd.merge(
            left=df,
            right=limitations_df,
            left_on='company',
            right_on='company',
            how='left')

        # add columns for counting unique emp|company values w/ totals
        df = BaseDataframe.training_per_emp(df)

        # check phone values
        df = BaseDataframe.phone_validation(df)

        # substring trainers from calendar via regex.
        df = BaseDataframe.trainer(df)

        # check if the training date is between the dates when company contract starts and ends
        df = BaseDataframe.active_contracts(df)

        # create a Month name column
        df["month"] = [pd.Timestamp(x).month_name() for x in df["start_time"]]

        # create a Year name column
        df["year"] = [pd.Timestamp(x).year for x in df["start_time"]]

        # add column with the name of the day when training was take place
        df['dayname'] = df['start_time'].dt.day_name()

        # reformat start_time
        df['training_datetime'] = BaseDataframe.datetime_normalize(df['start_time'])

        # change from datetime to date only
        df['scheduled_date'] = BaseDataframe.date_normalize(df['scheduled_on'])
        df['training_end'] = BaseDataframe.date_normalize(df['end_time'])

        # define the raw dataframes
        full_raw_report_df = df

        full_raw_report_df.attrs['name'] = "raw_full"

        monthly_raw_report_df = Transformation.annual_to_monthly_report_df(
            full_raw_report_df, Collection.datetime_final_format())

        monthly_raw_report_df.attrs['name'] = "raw_mont"

        # separate and select only needed columns for new pd sets
        columns_list = Collection.new_data_columns()

        # define the new dataframes
        new_full_data_df = full_raw_report_df[columns_list]

        new_monthly_data_df = monthly_raw_report_df[columns_list]

        # define the training session dataframes
        total_trainings_df, report_trainers_df = \
            BaseDataframe.total_trainings_func(new_monthly_data_df, new_full_data_df)

        dfs_dict = {
            "total_trainings_df": total_trainings_df,
            "report_trainers_df": report_trainers_df,
            "new_monthly_data_df": new_monthly_data_df,
            "new_full_data_df": new_full_data_df,
            "limitations_df": limitations_df,
            "flags_data_df": flags_data_df,
            "full_raw_report_df": full_raw_report_df,
            "monthly_raw_report_df": monthly_raw_report_df
        }

        return dfs_dict
