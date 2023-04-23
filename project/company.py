class Company:
    _limitation = 9999

    def __init__(self, name):
        self.name = name
        self.employees: List[Employee] = list()
        self.limitation = 9999
