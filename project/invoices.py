import os
from typing import List, Dict
from InvoiceGenerator.api import Invoice, Item, Client, Provider, Creator
from InvoiceGenerator.pdf import SimpleInvoice
from datetime import date


class BaseInvoice:

    @staticmethod
    def create_invoice(recipient: str, price: float, data_dict: dict):
        os.environ["INVOICE_LANG"] = "en"
        client = Client(recipient)
        provider = Provider('CouchMe Consultancy Services', bank_account='111-2222-33333', bank_code='BGSFXX')
        creator = Creator('CouchMe EOOD')

        invoice = Invoice(client, provider, creator)

        invoice.date = date.today()
        invoice.currency_locale = 'en_US.UTF-8'
        number_of_items = len(data_dict)

        for data in data_dict.items():
            nickname = data[0][0]
            service = data[0][1]
            units = data[1]
            # units = 1
            price_per_unit = price
            description = f"Employee: {nickname} for >>> '{service}'"
            invoice.add_item(Item(units, price_per_unit, description=description, tax=15))
        invoice.currency = "$"
        invoice.number = "10393069"
        document = SimpleInvoice(invoice)
        document.gen(f"exports/invoice_{recipient}.pdf") #generate_qr_code=True,


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
