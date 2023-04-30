from typing import List
import pandas as pd
from docxtpl import DocxTemplate


class ReportFromTemplate:

    @staticmethod
    def create_by_calendar_report_from_docx_template(name: str, df: pd.DataFrame, total_hours: float, total_pay: float):
        period = str(df['training_datetime'].iloc[0])[3:11]
        doc = DocxTemplate("imports/Reports_by_calendar.docx")
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
    def create_by_company_report_from_docx_template(company: str, df: pd.DataFrame, total_hours: float,
                                                    start_date, end_date):
        doc = DocxTemplate("imports/Reports_by_company.docx")
        row_list: List[list] = []
        df_columns = df.columns
        for i in df.index:
            row = [df[col][i] for col in df_columns]
            row_list.append(row)

        total_hours = int(total_hours)
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
    def create_by_company_x_report_from_docx_template(company: str, df: pd.DataFrame, total_hours: float,
                                                      start_date, end_date):
        doc = DocxTemplate("imports/Reports_by_company_x.docx")
        row_list: List[list] = []
        df_columns = df.columns
        for i in df.index:
            row = [df[col][i] for col in df_columns]
            row_list.append(row)

        total_hours = int(total_hours)
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
