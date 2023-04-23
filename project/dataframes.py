import pandas as pd


class DataframeTransformation:

    def __init__(self, name):
        self.name = name

    def nickname(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        The nickname must must have the following pattern:
        2 letters from the first name,
        3 letters from the last name and
        3 letters from the email (pvt or work w/o only alpha chars)
        """
        # adding column with both Fname and Lname
        df['employee_names'] = df['first_name'] + " " + df['last_name']

        # adding nickname column
        df['nickname'] = df['first_name'].str[0:2].str.upper() + df['last_name'].str[1:4].str.upper()

        # adding flag and 'X's at the end if the len < 5
        df.loc[df.nickname.str.len() < 5, 'flags'] += '1,'
        df.loc[df.nickname.str.len() < 5, 'nickname'] = df['nickname'].apply(lambda x: x + 'X' * (5 - len(x)))

        # check for email columns and take last part of the nickname
        # add flags for the cases
        work_email_validation_filter = (df['work_email'].str.contains('@', na=False))
        pvt_email_validation_filter = (df['pvt_email'].str.contains('@', na=False))

        df.loc[~work_email_validation_filter, 'flags'] += "2,"
        df.loc[~pvt_email_validation_filter, 'flags'] += "3,"

        df.loc[~pvt_email_validation_filter & ~work_email_validation_filter, 'flags'] += "4,"

        # takes value subtraction from the work or pvt email:
        email_validation_filter = (df['work_email'].str.contains('@', na=False))
        df.loc[email_validation_filter, 'nickname'] += \
            df['work_email'].str.replace("[^A-Za-z]+", "", regex=True).str[1:4].str.upper()

        df.loc[(df['work_email'].str.len() < 1) &
               (~df['pvt_email'].str.len() < 1) &
               (df['pvt_email'].str.contains('@')), 'nickname'] += \
            df['pvt_email'].str.replace("[^A-Za-z]+", "", regex=True).str[1:4].str.upper()

        df.loc[(df['work_email'].str.len() < 1) & (df['pvt_email'].str.len() < 1), 'nickname'] = \
            df['nickname'].apply(lambda x: x + 'X' * (8 - len(x)))

        df.loc[df['nickname'].str.len() < 8, 'nickname'] = df['nickname'].apply(lambda x: x + 'X' * (8 - len(x)))

        # adding column with both Fname and Lname
        df['employee_names'] = df['employee_names'].str.title()

        return df

    def company_substraction(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        1.Check if the meeting was held online or in-person and mark in a separate column
        2.Iterate through the 'type' column and extracts only the company name
        """

        df.loc[df['type'].str.contains('person|живо', regex=True), 'short_type'] = "На живо"
        df.loc[df['type'].str.contains('nline|нлайн', regex=True), 'short_type'] = "Онлайн"
        df.loc[~df['type'].str.contains('nline|нлайн|person|живо', regex=True), 'flags'] += "7,"

        df.loc[~df['type'].str.len() < 1, 'company'] = df['type'].str.split("[:|/]").str[0].str.upper().str.strip()
        return df

    def consultancy_per_emp(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Add three additional columns for:
            concatenation of employee|company,
            consultancy calculation total by the concat values,
            either the employee have one or more than one consultation
        """
        # concatenate nickname+company for later count and check if they are equal nicknames from different companies
        df.loc[~df['nickname'].str.len() < 1, 'concat_emp_company'] = df['nickname'].str[:] + "|" + df['company'].str[:]

        # count the frequency of occurrence for every  value (employee)
        df['total_per_emp'] = df.groupby('concat_emp_company')['concat_emp_company'].transform('count')

        # comment if the employee have one or more than one consultancies
        df.loc[df['total_per_emp'].astype(int) == 1, 'returns_or_not'] = 'only one session'

        df.loc[df['total_per_emp'].astype(int) > 1, 'returns_or_not'] = 'more then one session'
        return df

    def phone_validation(self, df: pd.DataFrame) -> pd.DataFrame:
        df.loc[df['phone'].str.len() < 1, 'flags'] += "5,"
        df.loc[df['phone'].str.contains('[^\d\+]', regex=True, na=False), 'flags'] += "6,"
        df.loc[df['phone'].str.len().between(1, 8), 'flags'] += "6,"
        return df

    def trainer(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Substract only the non cyrillic names and put them in a separate column
        """
        df.loc[~df['calendar'].isnull(), 'trainer'] = df['calendar'].str.split('|').str[0].str.strip().str.title()
        return df

    def active_contracts(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Count the consultancies of the company employees based on the consultancy
        date if there's an active contract'
        """
        # CONCAT the lines for which the consultancy date is between the contract's start and end date
        df.loc[(df['start_time'] >= df['starts']) & (df['start_time'] <= df['ends']), 'concat_count'] = \
            df['company'] + "|" + df['nickname']

        # count the company/employees total consultancies for the company active period
        df.loc[(df['start_time'] >= df['starts']) & (df['start_time'] <= df['ends']), 'active_cons_per_client'] = \
            df.groupby('concat_count')['concat_count'].transform('count')

        # calculate the number of consultancies that can be used
        df.loc[(df['start_time'] >= df['starts']) &
               (df['start_time'] <= df['ends']) &
               (df['c_per_emp'].between(1, 9998)), 'cons_left'] = \
            df['c_per_emp'] - df['active_cons_per_client']

        # add flag for employees which have less than 1 consultancy left
        df.loc[(~df['cons_left'].isna()) & (df['cons_left'] < 2), 'flags'] += '9,'
        return df

    def datetime_normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.apply(lambda x: x.strftime("%d-%b-%Y %H:%M:%S"))
        return df

    def date_normalize(self, df: pd.DataFrame) -> pd.DataFrame:
        df = pd.to_datetime(df).dt.date
        df = df.apply(lambda x: x.strftime("%d-%b-%Y"))
        return df

    def dates_diff(self, ser1: pd.Series, ser2: pd.Series) -> pd.Series:
        diff = ser2 - ser1
        return diff

    def total_consultations_func(self, df: pd.DataFrame) -> pd.DataFrame:
        df.insert(10, "language", 'Български', allow_duplicates=False)
        df.insert(10, "status", 'Проведен', allow_duplicates=False)
        return df

    def data_transformation(self, df: pd.DataFrame, df_limit) -> pd.DataFrame:
        """
        Transform the monthly and the annual df columns
        """
        df = df.rename(columns={
            'Start Time': 'start_time', 'End Time': 'end_time', 'Date Scheduled': 'scheduled_on',
            'First Name': 'f_name',
            'Last Name': 'l_name',
            'Phone': 'phone',
            'Email': 'pvt_email',
            'Type': 'type',
            'Calendar': 'calendar',
            'Appointment Price': 'price',
            'Paid?': 'is_paid',
            'Amount Paid Online': 'paid_online',
            'Certificate Code': 'certificate_code',
            'Notes': 'notes',
            'Label': 'label',
            'Scheduled By': 'scheduled_by',
            'Name of the company you work for | Име на фирмата, в която работите': 'company_name',
            'Work email | Служебен имейл': 'work_email',
            'Preferred platforms | Предпочитани платформи': 'preferred_platforms',
            'Appointment ID': 'appointment_id'
        })

        # add column for issues
        df['flags'] = ""
        df['emp_names_input'] = '|' + df['f_name'] + '|' + df['l_name'] + '|'

        # rough cleaning of the columns data
        df = unwanted_chars(df)
        df = trim_all_columns(df)

        # transliteration of cyrillic to latin chars
        df["first_name"] = transliteration(df, "f_name", "first_name")
        df["last_name"] = transliteration(df, "l_name", "last_name")

        # using first, last name and email parts
        df = nickname(df)

        # get only the company name and if the consultancy was IN PERSON/LIVE or ONLINE
        df = company_substraction(df)

        # merge/vlookup the columns from limitations to the monthly/annual df
        df = pd.merge(
            left=df,
            right=df_limit,
            left_on='company',
            right_on='company',
            how='left')

        # add columns for counting unique emp|company values w/ totals
        df = self.consultancy_per_emp(df)

        # check phone values
        df = self.phone_validation(df)

        # substring trainers from calendar via regex.
        df = self.trainer(df)

        # check if the consultancy date is between the dates when company contract starts and ends
        df = self.active_contracts(df)

        # create a Month name column
        df["month"] = [pd.Timestamp(x).month_name() for x in df["start_time"]]

        # create a Year name column
        df["year"] = [pd.Timestamp(x).year for x in df["start_time"]]

        # add column with the name of the day when consultancy was take place
        df['dayname'] = df['start_time'].dt.day_name()

        # reformat start_time
        df['cons_datetime'] = datetime_normalize(df['start_time'])

        # change from datetime to date only
        df['scheduled_date'] = date_normalize(df['scheduled_on'])
        df['cons_end'] = date_normalize(df['end_time'])

        return df


psy = DataframeTransformation("Psychea")
