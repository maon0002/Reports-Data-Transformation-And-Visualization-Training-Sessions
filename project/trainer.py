from typing import List

from project.appointment import Appointment
from project.company import Company
from project.employee import Employee


class Trainer:

    def __init__(self, names: str):
        self.names = names
        self.companies: List[Company] = list()
        self.clients: List[Employee] = list()
        self.appointments: List[Appointment] = list()
