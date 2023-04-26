from typing import List, Dict
import pandas as pd


class Collection:

    @staticmethod
    def date_default_format():
        return "%d-%b-%Y"

    @staticmethod
    def datetime_default_format():
        return "%d-%b-%Y %H:%M:%S"

    @staticmethod
    def datetime_final_format():
        return "%Y-%m-%d %H:%M:%S"

    @staticmethod
    def flags_dict() -> dict:
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
    def transliterate_bg_to_en(df: pd.DataFrame, column: str, new_column: str) -> pd.Series:

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

        # transliteration char by char
        transliterated_string = ""
        for i in range(len(df[column])):
            string_value = df.loc[i, column].strip()
            if string_value and not string_value.isascii():
                for char in string_value:
                    if char in bg_en_dict.keys():
                        string_value = string_value.replace(
                            char, bg_en_dict[char])
                        transliterated_string = string_value
                    else:
                        transliterated_string = string_value

                df.loc[i, new_column] = transliterated_string.upper()
                transliterated_string = ""
            else:
                df.loc[i, new_column] = df.loc[i, column].upper()
        return df[new_column]

    @staticmethod
    def new_data_columns() -> List[str]:
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
    def trainers_columns() -> List[str]:
        return [
            'type',
            'trainer',
            'company',
            'employee_names',
            'training_datetime',
            'short_type',
            'status'
        ]

    @staticmethod
    def raw_report_list() -> List[str]:
        return ['new_monthly_data_df', 'new_full_data_df', 'limitations', 'flags_data']

    @staticmethod
    def generic_report_list() -> List[str]:
        return [
            "total_trainings_df",
            "report_trainers_df",
            "new_monthly_data_df",
            "new_full_data_df",
            "limitations_df",
            "flags_data_df"]

    @staticmethod
    def company_report_list(new_monthly_data_df) -> List[str]:
        # list the companies with is_valid == 1
        return [comp for comp in new_monthly_data_df[(new_monthly_data_df['is_valid'] == 1)]['company'].unique()]

    @staticmethod
    def company_report_list_other(new_monthly_data_df) -> List[str]:
        # list the companies with is_valid == 1
        return [comp_other for comp_other in new_monthly_data_df[(new_monthly_data_df['is_valid'] == 0)]['company']
                .unique()]

    @staticmethod
    def trainers_report_list(monthly_data: pd.DataFrame) -> List[str]:
        return [trainer for trainer in monthly_data['trainer'].unique()]
