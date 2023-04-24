from typing import List
from project.trainers import Trainer


class Employee:

    def __init__(self, names: str, phone: str, pvt_email: str, corp_email: str, company: str):
        self.names = names
        self.phone = phone
        self.pvt_email = pvt_email
        self.corp_email = corp_email
        self.company = company
        self.trainers: List[Trainer] = list()
        self.appointments = 0
