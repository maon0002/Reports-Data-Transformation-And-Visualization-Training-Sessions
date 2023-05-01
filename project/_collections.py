from typing import List, Dict
import pandas as pd



class Collection:
    """Class used to keep a different collections

        Attributes
        ----------
        No attributes

        Methods
        -------
        report_expected_columns():
            :return: the expected names and order of the imported Initial/General report

        limitations_expected_columns():
            :return: the expected names and order of the imported Limitations report

        date_default_format():
            :return: a date format (09-Feb-2023)

        datetime_default_format():
            :return: a datetime format (23-Dec-2020 16:00:00)

        datetime_final_format():
            :return: a datetime format (2020-12-23 16:00:00)

        flags_dict() -> dict:
            :return: dictionary with flag numbers and flags meanings

        transliterate_dict() -> Dict[str, str]:
            :return: dictionary with every bg letter as key and the latin equivalent as value

        new_data_columns() -> List[str]:
            :return: series/columns for the new dataframes

        trainings_columns() -> List[List[str]]:
            :return: two lists with series/column names for the new dataframes

        generic_report_list() -> List[str]:
            :return: list with dataframe names

        company_report_list(new_monthly_data_df) -> List[str]:
            :return: list the companies with is_valid == 1

        company_report_list_other(new_monthly_data_df) -> List[str]:
            :return: list the companies with is_valid == 0

        trainers_report_list(monthly_data: pd.DataFrame) -> List[str]:
            :return: list with trainer names

    """

    @staticmethod
    def report_expected_columns():
        """
        :return: the expected names and order of the imported Initial/General report
        """
        return ['Start Time', 'End Time', 'First Name', 'Last Name', 'Phone', 'Email', 'Type', 'Calendar',
                'Appointment Price', 'Paid?', 'Amount Paid Online', 'Certificate Code', 'Notes',
                'Date Scheduled', 'Label', 'Scheduled By',
                'Име на компанията, в която работите | Name of the company you work for  ',
                'Служебен имейл | Work email  ', 'Предпочитани платформи | Preferred platforms  ',
                'Appointment ID']

    @staticmethod
    def limitations_expected_columns():
        """
        :return: the expected names and order of the imported Limitations report
        """
        return ['COMPANY', 'C_PER_PERSON', 'C_PER_MONTH', 'PREPAID', 'START', 'END', 'DURATION DAYS',
                'NOTE', 'BGN_PER_HOUR', 'IS_VALID']

    @staticmethod
    def date_default_format():
        """
        :return: a date format (09-Feb-2023)
        """
        return "%d-%b-%Y"

    @staticmethod
    def datetime_default_format():
        """
        :return: a datetime format (23-Dec-2020 16:00:00)
        """
        return "%d-%b-%Y %H:%M:%S"

    @staticmethod
    def datetime_final_format():
        """
        :return: a datetime format (2020-12-23 16:00:00)
        """
        return "%Y-%m-%d %H:%M:%S"

    @staticmethod
    def flags_dict() -> dict:
        """
        :return: dictionary with flag numbers and flags meanings
        """
        return {
            'flag_number': [1, 2, 3, 4, 5, 6, 7, 8, 9],
            'flag_note':
                ['names issue',
                 'work_mail issue',
                 'pvt_mail issue',
                 'missing both pvt_email and work_email',
                 'missing phone number',
                 'phone number issue',
                 'missing short_type',
                 'pvt email equal to work email',
                 'number of the trainings left less than 1'
                 ]}

    @staticmethod
    def transliterate_dict() -> Dict[str, str]:
        """
        :return: dictionary with every bg letter as key and the latin equivalent as value
        """
        """
        Transliterate bg chars into latin when needed
        """
        # create dictionary for transliteration from bg to latin letters
        bg_en_dict = {
            'А': 'A',
            'Б': 'B',
            'В': 'V',
            'Г': 'G',
            'Д': 'D',
            'Е': 'E',
            'Ж': 'Zh',
            'З': 'Z',
            'И': 'I',
            'Й': 'Y',
            'К': 'K',
            'Л': 'L',
            'М': 'M',
            'Н': 'N',
            'О': 'O',
            'П': 'P',
            'Р': 'R',
            'С': 'S',
            'Т': 'T',
            'У': 'U',
            'Ф': 'F',
            'Х': 'H',
            'Ц': 'Ts',
            'Ч': 'Ch',
            'Ш': 'Sh',
            'Щ': 'Sht',
            'Ъ': 'A',
            'Ь': 'Y',
            'Ю': 'Yu',
            'Я': 'Ya',
            'а': 'a',
            'б': 'b',
            'в': 'v',
            'г': 'g',
            'д': 'd',
            'е': 'e',
            'ж': 'zh',
            'з': 'z',
            'и': 'i',
            'й': 'y',
            'к': 'k',
            'л': 'l',
            'м': 'm',
            'н': 'n',
            'о': 'o',
            'п': 'p',
            'р': 'r',
            'с': 's',
            'т': 't',
            'у': 'u',
            'ф': 'f',
            'х': 'h',
            'ц': 'ts',
            'ч': 'ch',
            'ш': 'sh',
            'щ': 'sht',
            'ъ': 'a',
            'ь': 'y',
            'ю': 'yu',
            'я': 'ya',
            ' ': ' ',
            '-': '-'}

        return bg_en_dict

    @staticmethod
    def new_data_columns() -> List[str]:
        """
        :return: series/columns for the new dataframes
        """
        return [
            'training_datetime',
            'dayname',
            'month',
            'year',
            'nickname',
            'employee_names',
            'company',
            'trainer',
            'short_type',
            'bgn_per_hour',
            'concat_emp_company',
            'total_per_emp',
            'trainings_left',
            'returns_or_not',
            'flags',
            'phone',
            'pvt_email',
            'work_email',
            'scheduled_date',
            'type',
            'emp_names_input',
            'is_valid'
        ]

    @staticmethod
    def trainings_columns() -> List[List[str]]:
        """
        :return: two lists with series/column names for the new dataframes
        """
        return [[
            'concat_emp_company',
            'type',
            'company',
            'nickname',
            'training_datetime',
            'employee_names',
            'work_email',
            'trainer',
            'short_type'
        ], [
            'type',
            'trainer',
            'company',
            "employee_names",
            'training_datetime',
            'work_email',
            'short_type',
            'status'
        ]]

    @staticmethod
    def generic_report_list() -> List[str]:
        """
        :return: list with dataframe names
        """
        return [
            "total_trainings_df",
            "report_trainers_df",
            "new_monthly_data_df",
            "new_full_data_df",
            "limitations_df",
            "flags_data_df"]

    @staticmethod
    def company_report_list(new_monthly_data_df) -> List[str]:
        """
        :param new_monthly_data_df:
        :return: list the companies with is_valid == 1
        """
        return [comp for comp in new_monthly_data_df[(new_monthly_data_df['is_valid'] == 1)]['company'].unique()]

    @staticmethod
    def company_report_list_other(new_monthly_data_df) -> List[str]:
        """
        :param new_monthly_data_df:
        :return: list the companies with is_valid == 0
        """
        return [comp_other for comp_other in
                new_monthly_data_df[(new_monthly_data_df['is_valid'] == 0)]['company'].unique()]

    @staticmethod
    def trainers_report_list(monthly_data: pd.DataFrame) -> List[str]:
        """
        :param monthly_data:
        :return: list with trainer names
        """
        return [trainer for trainer in monthly_data['trainer'].unique()]


