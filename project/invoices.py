import os
from typing import List, Dict
from InvoiceGenerator.api import Invoice, Item, Client, Provider, Creator
from InvoiceGenerator.pdf import SimpleInvoice
from datetime import date


class BaseInvoice:
    __invoice_next_seq_file = "invoice_next_seq.txt"

    @staticmethod
    def get_invoice_number():
        seq_file = f"imports/{BaseInvoice.__invoice_next_seq_file}"
        invoice_current_number = 0
        f = open(seq_file, "r+")
        value = f.readline()
        invoice_current_number = value
        new_number = int(invoice_current_number) + 1
        f.truncate(0)
        f.close()
        BaseInvoice.update_invoice_number(seq_file, new_number)
        return invoice_current_number

    @staticmethod
    def update_invoice_number(path, new_number):
        f = open(path, "r+")
        f.write(str(new_number))
        f.close()

    @staticmethod
    def create_invoice(recipient: str, price: float, data_dict: dict):
        os.environ["INVOICE_LANG"] = "en"
        client = Client(recipient)
        provider = Provider(summary='CouchMe', address='Couching str. 55A', zip_code='1000', city='Indore',
                            country='Bulgaria', bank_name='KBC Bank', bank_code='BGSFKBC',
                            bank_account='KBC000SF999888',
                            phone='+35987654321', email='couchme@mail.com', logo_filename='images/logo.webp')
        creator = Creator('CouchMe EOOD')
        invoice = Invoice(client, provider, creator)
        invoice.date = date.today()
        invoice.currency_locale = 'bg_BG.UTF-8'
        invoice.currency = 'BGN'
        invoice.logo_filename = "project/images/logo.png"
        invoice.number = BaseInvoice.get_invoice_number()

        for data in data_dict.items():
            nickname = data[0][0]
            service = data[0][1]
            units = data[1]
            price_per_unit = price
            description = f"Employee: {nickname} for >>> '{service}'"
            invoice.add_item(Item(units, price_per_unit, description=description, tax=15))

        document = SimpleInvoice(invoice)
        document.gen(f"exports/invoice_{recipient}.pdf")


#
# test = BaseInvoice()
#
# test.create_invoice("MaonCorp", 55.00,
#                     {('AGSHLASH', 'SunriseCorp: Тренинг за лидери на живо | Leadership training in person'): 2,
#                      ('ANONTMON', 'SunriseCorp: Онлайн тренинг за лидери | Online leadership training'): 2,
#                      ('SEARATAR', 'SunriseCorp: Тренинг за лидери на живо | Leadership training in person'): 2,
#                      ('AYONRCON', 'SunriseCorp: Онлайн тренинг за лидери | Online leadership training'): 1,
#                      ('DEATAKAT', 'SunriseCorp: Тренинг за лидери на живо | Leadership training in person'): 1,
#                      ('DIELEBEL', 'SunriseCorp: Онлайн тренинг за лидери | Online leadership training'): 1,
#                      ('GEHUPCHU', 'SunriseCorp: Онлайн тренинг за лидери | Online leadership training'): 1,
#                      ('KAAMIRAM', 'SunriseCorp: Тренинг за лидери на живо | Leadership training in person'): 1,
#                      ('MAARNARI', 'SunriseCorp: Онлайн тренинг за лидери | Online leadership training'): 1,
#                      ('PAANGAVI', 'SunriseCorp: Онлайн тренинг за лидери | Online leadership training'): 1,
#                      ('PEANGMAN', 'SunriseCorp: Тренинг за лидери на живо | Leadership training in person'): 1,
#                      ('PERUSKRU', 'SunriseCorp: Тренинг за лидери на живо | Leadership training in person'): 1,
#                      ('SEARVPAR', 'SunriseCorp: Тренинг за лидери на живо | Leadership training in person'): 1})
