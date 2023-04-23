import pandas as pd


class Import:

    @staticmethod
    def to_lower(x): return x.lower() if isinstance(x, str) else x

    @staticmethod
    def import_report(path) -> pd.DataFrame:
        report_df = pd.read_csv(path,
                                parse_dates=['Start Time',
                                             'End Time', 'Date Scheduled'],
                                infer_datetime_format=True,
                                dayfirst=True,
                                cache_dates=True,
                                dtype={
                                    'First Name': 'string',
                                    'Last Name': 'string',
                                    'Phone': 'string',
                                    'Type': 'string',
                                    'Calendar': 'string',
                                    'Appointment Price': 'float',
                                    'Paid?': 'string',
                                    'Amount Paid Online': 'float',
                                    'Certificate Code': 'string',
                                    'Notes': 'string',
                                    'Label': 'string',
                                    'Scheduled By': 'string',
                                    'Name of the company you work for | Име на фирмата, в която работите': 'string',
                                    'Preferred platforms | Предпочитани платформи': 'string',
                                    'Appointment ID': 'string'},
                                converters={
                                    'Email': Import.to_lower, 'Work email | Служебен имейл': Import.to_lower}, ) \
            .fillna(np.nan).replace([np.nan], [None])
        return report_df

    @staticmethod
    def import_limitations(path):
        limitations_df = pd.read_csv(path,
                                     names=['company', 'c_per_emp', 'c_per_month', 'prepaid',
                                            'starts', 'ends', 'note', 'bgn_per_hour', 'is_valid'],
                                     usecols=[0, 1, 2, 3, 4, 5, 7, 8, 9],
                                     parse_dates=['starts', 'ends'],
                                     infer_datetime_format=True,
                                     dayfirst=True,
                                     cache_dates=True,
                                     skiprows=[0],
                                     index_col=False,
                                     keep_default_na=False,
                                     )
        return limitations_df


class Transform:

    @staticmethod
    def annual_to_monthly_report_df(report_dataframe: pd.DataFrame) -> pd.DataFrame:
        last_row_value = (report_dataframe['Start Time'].tail(1).dt.to_period('M'))
        monthly_data_df = report_dataframe.loc[report_dataframe['Start Time'].dt.to_period('M').isin(last_row_value)]. \
            reset_index(drop=True)
        return monthly_data_df


class Export:
    pass


class Update:
    pass


class Create:
    pass
