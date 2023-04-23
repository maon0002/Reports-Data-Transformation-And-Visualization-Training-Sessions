import pandas as pd


class BaseDataframe:

    def __init__(self, name: str, dataframe: pd.DataFrame):
        self.name = name
        self.dataframe = dataframe
