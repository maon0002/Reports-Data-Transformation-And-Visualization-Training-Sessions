import os
from InvoiceGenerator.api import Invoice, Item, Client, Provider, Creator
from InvoiceGenerator.pdf import SimpleInvoice
from datetime import date


class BaseInvoice:
    @staticmethod
    def invoice_seq_file():
        return "imports/invoice_next_seq.txt"

    @staticmethod
    def invoice_filepath():
        return "exports/invoices/"

    @staticmethod
    def get_invoice_number() -> str:
        """
        It gets a number from a .txt file, adds 1 to be ready for the next invoice seq number
        :return: seq number as a string to be used as invoice number (for example 1111111112)
        """
        seq_file = BaseInvoice.invoice_seq_file()
        f = open(seq_file, "r+")
        value = f.readline()
        invoice_current_number = value
        new_number = int(invoice_current_number) + 1
        f.truncate(0)
        f.close()
        BaseInvoice.update_invoice_number(seq_file, new_number)
        return invoice_current_number

    @staticmethod
    def update_invoice_number(path, new_number) -> None:
        """
        Clear the file with the previous invoice number and replaced it with a new number (old number + 1)
        :param path: path to the file which contains one number, the number of the current invoice
        :param new_number: new number with which the old number from the .txt must be replaced (for example 1111111113)
        :return: nothing
        """
        f = open(path, "r+")
        f.write(str(new_number))
        f.close()

    @staticmethod
    def create_invoice(recipient: str, price: float, data_dict: dict) -> None:
        """
        During the data aggregation for the .xlsx reports for each company, this function generates an invoice on a
        employee/service level
        :param recipient: company name
        :param price: the price for the company based on the contract with the service provider
        :param data_dict: the data from the dataframe, converted into dictionary in the following format:
        {('NICKNAME', 'COMPANY:Description in Bulgarian | Description in English'): quantity(int)}
        example: {('TORADGRA', 'QuantumPeak:Тренинг за лидери на живо | Leadership training in person'): 1}
        :return: nothing
        """
        os.environ["INVOICE_LANG"] = "en"
        client = Client(recipient)
        provider = Provider(summary='CouchMe', address='Couching str. 55A', zip_code='1000', city='Sofia',
                            country='Bulgaria', bank_name='KBC Bank', bank_code='BGSFKBC',
                            bank_account='KBC000SF999888',
                            phone='+35987654321', email='couchme@mail.com', logo_filename='images/logo.webp')
        creator = Creator('CouchMe EOOD')
        invoice = Invoice(client, provider, creator)
        invoice.date = date.today()
        invoice.currency_locale = 'bg_BG.UTF-8'
        invoice.currency = 'BGN'

        invoice.number = BaseInvoice.get_invoice_number()

        # or use manual input
        # invoice.number = 0
        # while True:
        #     invoice_number = input("Please enter the initial invoice number: ")
        #     invoice_number = invoice_number.strip()
        #     if invoice_number.isdigit() and len(invoice_number) == 10:
        #         invoice.number = invoice_number
        #         break

        invoice.use_tax = True
        for data in data_dict.items():
            nickname = data[0][0]
            service = data[0][1]
            units = data[1]
            price_per_unit = price
            description = f"Employee: {nickname} for >>> '{service}'"
            invoice.add_item(Item(units, price_per_unit, description=description, tax=15))
        document = SimpleInvoice(invoice)
        invoice_path = BaseInvoice.invoice_filepath()
        document.gen(f"{invoice_path}invoice_{recipient}.pdf")
