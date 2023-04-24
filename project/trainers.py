from typing import List

from project.appointments import Appointment
from project.companies import Company
from project.employees import Employee


class Trainer:

    def __init__(self, names: str):
        self.names = names
        self.companies: List[Company] = list()
        self.clients: List[Employee] = list()
        self.appointments: List[Appointment] = list()
