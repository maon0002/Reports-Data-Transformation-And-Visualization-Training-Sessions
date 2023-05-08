from typing import List
import pandas as pd
from docxtpl import DocxTemplate


class ReportFromTemplate:
    """Class used to represent a report based on pre-build/pre-formatted
        .docx template.
        By idea it serves the BulkReport report class and creates separate
        reports for each company or trainer filtered in the new monthly dataframe.

        Attributes
        ----------
        No attributes

        Methods
        -------
        create_by_calendar_report_from_docx_template(name: str, df: pd.DataFrame, total_hours: float, total_pay: float
                                                    ) -> None:

            Render and save a separate .docx report for every trainer by his/her name

        create_by_company_report_from_docx_template(company: str, df: pd.DataFrame, total_hours: float, start_date,
                                                    end_date ) -> None:

            Render and save a separate .docx report for every company by name and
            is_valid column (excludes out of scope services).
            *with more details compared with the by_company_x template

        create_by_company_x_report_from_docx_template(company: str, df: pd.DataFrame, total_hours: float,
                                                      start_date, end_date):

            Render and save a separate .docx report for every company by name and
            is_valid column (excludes out of scope services)
            *Variation of by_company template
            **with less details compared with the by_company_x template
        """
    @staticmethod
    def create_by_calendar_report_from_docx_template(name: str,
                                                     df: pd.DataFrame,
                                                     total_hours: float,
                                                     total_pay: float
                                                     ) -> None:
        """
        The function gets filtered and aggregated data from the
        'new_monthly_data_df' for every trainer,
        fill the data into a cells with '{{...}}}' values
        in a .docx pre-formatted template
        and save the file

        :param name: The name of the trainer
        :param df: The dataframe records filtered by the name of the trainer
        :param total_hours: The sum of the total hours/trainings for the particular trainer
        :param total_pay: total_hours multiplied by bgn_per_hour dataframe column value (depends on company contract)
        :return: None
        """
        # get the Month-Year report period
        period = str(df['training_datetime'].iloc[0])[3:11]

        # Make variable with the template object with the path to the .docx template
        doc = DocxTemplate("imports/Reports_by_calendar.docx")

        # row_list combined with the {{...}} values from the template
        # are the magic of the looping trough the records and
        # replacing/adding them in the template
        row_list: List[list] = []
        df_columns = df.columns
        for i in df.index:
            row = [df[col][i] for col in df_columns]
            row_list.append(row)

        total_hours = int(total_hours)
        total_hours_sum = f"{total_pay:.2f}"
        minus_description = ""
        minus_value = 0
        to_receive = f"{(float(total_hours_sum) - minus_value):.2f}"

        doc.render({
            "calendar": name,
            "period": period,
            "invoice_list": row_list,
            "status": "Проведен",
            "total_hours": total_hours,
            "total_hours_sum": total_hours_sum,
            "minus_description": minus_description,
            "minus_value": minus_value,
            "to_receive": to_receive,
        })
        doc.save(f"exports/from_templates/by_calendar/Reports_by_calendar_{name}.docx")

    @staticmethod
    def create_by_company_report_from_docx_template(company: str,
                                                    df: pd.DataFrame,
                                                    total_hours: float,
                                                    start_date, end_date
                                                    ) -> None:
        """
        The function gets filtered and aggregated data from the
        'new_monthly_data_df' for every company,
        fill the data into a cells with '{{...}}}' values
        in a .docx pre-formatted template
        and save the file

        :param company: The name of the company
        :param df: The dataframe records filtered by the name of the company
        :param total_hours: The sum of the total hours/trainings for all of the company employees
        :param start_date: the first available date of employee training for the current month
        :param end_date: the last available date of employee training for the current month
        :return: None
        """

        # Make variable with the template object with the path to the .docx template
        doc = DocxTemplate("imports/Reports_by_company.docx")

        # row_list combined with the {{...}} values from the template
        # are the magic of the looping trough the records and
        # replacing/adding them in the template
        row_list: List[list] = []
        df_columns = df.columns
        for i in df.index:
            row = [df[col][i] for col in df_columns]
            row_list.append(row)

        total_hours = int(total_hours)

        # empty strings can be used in future
        prepaid = ""
        remaining = ""
        upcoming = ""
        default = 0
        total_used = total_hours

        doc.render({

            "company": company,
            "period_start": start_date,
            "period_end": end_date,
            "invoice_list": row_list,
            "total_hours": total_hours,
            "prepaid": prepaid,
            "remaining": remaining,
            "default": default,
            "upcoming": upcoming,
            "total_used": total_used,
        })
        doc.save(f"exports/from_templates/by_company/Reports_by_company_{company}.docx")

    @staticmethod
    def create_by_company_x_report_from_docx_template(company: str,
                                                      df: pd.DataFrame,
                                                      total_hours: float,
                                                      start_date, end_date):
        """
        The function gets filtered and aggregated data from the
        'new_monthly_data_df' for every company,
        fill the data into a cells with '{{...}}}' values
        in a .docx pre-formatted template
        and save the file

        :param company: The name of the company
        :param df: The dataframe records filtered by the name of the company
        :param total_hours: The sum of the total hours/trainings for all of the company employees
        :param start_date: the first available date of employee training for the current month
        :param end_date: the last available date of employee training for the current month
        :return: None
        """

        # Make variable with the template object with the path to the .docx template
        doc = DocxTemplate("imports/Reports_by_company_x.docx")

        # row_list combined with the {{...}} values from the template
        # are the magic of the looping trough the records and
        # replacing/adding them in the template
        row_list: List[list] = []
        df_columns = df.columns
        for i in df.index:
            row = [df[col][i] for col in df_columns]
            row_list.append(row)

        total_hours = int(total_hours)
        # empty string can be used in future
        upcoming = ""
        default = 0
        total_used = total_hours

        doc.render({
            "company": company,
            "period_start": start_date,
            "period_end": end_date,
            "invoice_list": row_list,
            "total_hours": total_hours,
            "default": default,
            "upcoming": upcoming,
            "total_used": total_used,
        })
        doc.save(f"exports/from_templates/by_company_x/Reports_by_company_x_{company}.docx")
