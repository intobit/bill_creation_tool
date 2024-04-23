import fitz
import re


class PdfReader:
    def __init__(self, path_to_pdf: str):
        self.file = path_to_pdf
        self.open_file = fitz.open(self.file)

    def get_invoice_and_delivery_data(self):
        # Invoice address data
        invoice_data = {
            "c_name": '',
            "c_owner_name": '',
            "c_street": '',
            "c_street_number": '',
            "c_zip_code": '',
            "c_city": '',
            "c_country": '',
            "c_tax_num": ''
        }
        invoice_textboxes = ["Text Box 1", "Text Box 2", "Text Box 3",
                             "Text Box 4", "Text Box 5", "Text Box 6",
                             "Text Box 7", "Text Box 8"]

        # Delivery address data
        delivery_data = {
            "d_name": '',
            "d_street": '',
            "d_street_number": '',
            "d_zip_code": '',
            "d_city": '',
            "d_country": '',
        }

        delivery_textboxes = ["Text Box 9", "Text Box 10", "Text Box 11",
                              "Text Box 12", "Text Box 13", "Text Box 14"]

        invoice_index = 0
        invoice_keys = list(invoice_data.keys())
        i_keys_index = 0

        delivery_index = 0
        delivery_keys = list(delivery_data.keys())
        d_keys_index = 0

        for page in self.open_file:
            for widget in page.widgets():
                field_name = widget.field_name
                field_value = widget.field_value

                if field_name in invoice_textboxes:
                    i_key = invoice_keys[i_keys_index]
                    invoice_data[i_key] = field_value
                    invoice_index += 1
                    i_keys_index += 1

                    if invoice_index == 13:
                        invoice_data[invoice_keys[i_keys_index]] = widget.next.field_value

                if field_name in delivery_textboxes:
                    d_key = delivery_keys[d_keys_index]
                    delivery_data[d_key] = field_value
                    d_keys_index += 1
                    delivery_index += 1
                else:
                    invoice_index += 1
                    delivery_index += 1
        return invoice_data, delivery_data

    def get_product_orders(self):
        orders = {}

        for page_number in range(len(self.open_file)):
            page = self.open_file.load_page(page_number)

            table = page.find_tables()
            for line in table[0].extract():
                line_str = ""
                for item in line:
                    if item is not None:
                        line_str += item
                    if re.findall('[0-9]', line_str) and line[1] != '':
                        orders[line[0]] = line[1]
        return orders

    def total_price(self, price: float = 4.2):
        products = self.get_product_orders()
        total = 0
        price = price
        for value in products.values():
            total += int(value)
        final_price = round(total * price, 2)
        return final_price

