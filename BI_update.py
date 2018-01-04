import os
import argparse

import xlrd
from openpyxl import Workbook


SKUs = {
    'CHHC': 'Chardonnay HC',
    'CHZV': 'Chardonnay',
    'CSHC': 'Cabernet Sauvignon HC',
    'GBZV': 'Grenache Blanc',
    'GRZV': 'Grenache',
    'MOZV': 'Mourvedre',
    'PBBNV': 'Pinot Blanc HC',
    'PBHC': 'Pinot Blanc HC',
    'PNHC': 'Pinot Noir HC',
    'ROZV': 'Roussanne',
    'SBMV': 'Sauvignon Blanc HC',
    'SY30TH': 'Syrah 30th Annv. 1.5L',
    'SY8BBL': 'Eight Barrel Syrah',
    'SYAMP': 'Syrah Amphora',
    'SYB3': 'Black Bear Syrah',
    'SYCHG': 'Chapel G Syrah',
    'SYESTRELLA': 'Estrella Syrah',
    'SYM1.5L': 'Mesa Reserve Syrah 1.5L',
    'SYMESA': 'Mesa Reserve Syrah',
    'SYMESA1.5L': 'Mesa Reserve Syrah 1.5L',
    'SYMESA15': 'Mesa Reserve Syrah 1.5L',
    'SYZV': 'Syrah',
    'SYZV1.5L': 'Syrah 1.5L',
    'SYZV3L': 'Syrah 3L',
    'SYZV5L': 'Syrah 5L',
    'TOY': 'Toyon',
    'VIZV': 'Viognier',
    'Z3': 'Z Three',
    'Z31.5L': 'Z Three 1.5L',
    'Z3ZV': 'Z Three',
    'ZBZV': 'Z Blanc',
    'ZCZV': 'Z Cuvee',
    'ZGZV': 'Z Gris Rose'
}


def parse_arguments():
    """
    Sets up the terminal arguments.
    :return: tuple – the input and output filenames
    """
    parser = argparse.ArgumentParser(description='Process the Vintegrate data for Power BI.',
                                     epilog='NB: Place this file in a direct sub-folder of the one that '
                                     'contains the data.')
    parser.add_argument('item_sales_output_filename',
                        help='The output filename for itemsales.xls. Omit the \'.xls\'')
    parser.add_argument('invoice_details_output_filename',
                        help='The output filename for invoice_details.xls. Omit the \'.xls\'')
    parser.add_argument('--salesinput',
                        default='itemsales.xls',
                        dest='sales_input_filename',
                        help='Specify this value if different from default: itemsales.xls')
    parser.add_argument('--invoicesinput',
                        default='directmarketingreporttransactionitem.xls',
                        dest='invoices_input_filename',
                        help='Specify this value if different from default: directmarketingreporttransactionitem.xls')

    args = parser.parse_args()
    return (args.sales_input_filename,
            args.invoices_input_filename,
            args.item_sales_output_filename,
            args.invoice_details_output_filename)


def get_data(filename):
    """
    Reads data in from .xls file using `xlrd`.
    :param filename: The file to open. Format: `<filename>.xls`
    :return: a list of lists `[[row]]`
    """
    path = '../' + filename
    try:
        book = xlrd.open_workbook(path)
    except FileNotFoundError as e:
        print("Your input file `{}` is either missing or not in the right location.".format(e.filename))
        raise
    sheet = book.sheets()[0]

    # Convert all values into native Python types as a list of lists [[rows]]
    data = []
    for i in range(sheet.nrows):
        data.append(sheet.row_values(rowx=i))

    return data


def save(data, filename):
    """
    Saves data to a .xlsx file via `openpyxl`.
    :param data: A list of lists, [[row]], consisting of the processed data.
    :param filename: The file to write to. Format as `<filename>` (no extension)
    :return: None
    """
    wb = Workbook()
    ws = wb.active

    for row in data:
        ws.append(row)

    path = '../' + filename + '.xlsx'
    wb.save(path)


def process_item_sales(sales_input_filename, item_sales_output_filename):
    """
    Converts raw item sales data into usable format for BI.
    :param sales_input_filename: The file to open. Format: `<filename>.xls`
    :param item_sales_output_filename: The file to write to. Format as `<filename>` (no extension)
    :return: None
    """
    cells = get_data(sales_input_filename)

    # Reformat header to: ['Last Name', 'First Name', 'Order Number', 'Invoice Date', 'Quantity']
    cells[0] = cells[0][12:]

    # Insert rows to hold SKUs and Product Names and remove excess rows
    for i in range(len(cells)):
        cells[i].insert(0, None)
        cells[i].insert(0, None)
        cells[i] = cells[i][:7]

    cells[0][0] = 'Varietal'
    cells[0][1] = 'Product'

    # Get the SKUs and Product Names in all rows
    for i in range(1, len(cells)):
        if cells[i][4] == '':
            cells[i][0] = cells[i][2]
            cells[i][1] = cells[i][3]
        else:
            cells[i][0] = cells[i - 1][0]
            cells[i][1] = cells[i - 1][1]

    # Filter out the non-wines
    wine_rows = [x for x in cells[1:] if x[1][:2] == '20']
    wine_rows.insert(0, cells[0])

    # Remove the YR from the SKUs (e.g. 10SYZV3L -> SYZV3L)
    for row in wine_rows[1:]:
        row[0] = row[0][2:]

    # Remove the SKU category headers and sums (interspersed throughout rows)
    purchases = [x for x in wine_rows if x[6] != '']

    # Convert SKUs to Varietals
    for row in purchases[1:]:
        try:
            row[0] = SKUs[row[0]]
        except KeyError:
            print("Need to add an SKU to the dictionary: ", row[0])
            raise

    # Insert Full Name field
    purchases[0].insert(2, 'Full Name')
    for row in purchases[1:]:
        row.insert(2, row[2] + ", " + row[3])

    # Reformat "Order Number" and "Quantity" to int
    for row in purchases[1:]:
        row[5] = int(row[5])
        row[7] = int(row[7])

    # Write to Excel file and delete original file
    save(purchases, item_sales_output_filename)


def process_invoice_details(invoice_input_filename, invoice_details_output_filename):
    """
    Converts raw invoice details data into usable format for BI.
    :param invoice_input_filename: The file to open. Format: `<filename>.xls`
    :param invoice_details_output_filename: The file to write to. Format as `<filename>` (no extension)
    :return: None
    """
    data = get_data(invoice_input_filename)

    # Removing 'Sales Rep' and 'Payment Type' fields
    cropped_data = [row[:11] for row in data]

    # Filtering to only have invoices of wine sales.
    invoices = [row for row in cropped_data if row[0] != '']
    wine_invoices = [row for row in invoices[1:] if row[5][:2] == '20']
    wine_invoices.insert(0, invoices[0])

    # Remove the YR from the SKUs (e.g. 10SYZV3L -> SYZV3L)
    for row in wine_invoices[1:]:
        row[4] = row[4][2:]

    # Convert SKUs to varietals
    for row in wine_invoices[1:]:
        try:
            row[4] = SKUs[row[4]]
        except KeyError:
            print("Need to add an SKU to the dictionary: ", row[4])
            raise
    wine_invoices[0][4] = 'Varietal'

    # Reformat 'Invoice Number' and 'Quantity' to int
    for row in wine_invoices[1:]:
        row[0] = int(row[0])
        row[6] = int(row[6])

    save(wine_invoices, invoice_details_output_filename)


def main(sales_input_filename, invoice_input_filename, item_sales_output_filename, invoice_details_output_filename):
    process_item_sales(sales_input_filename, item_sales_output_filename)
    process_invoice_details(invoice_input_filename, invoice_details_output_filename)

    # Remove source data files.
    os.remove('../' + sales_input_filename)
    os.remove('../' + invoice_input_filename)


if __name__ == '__main__':
    arguments = parse_arguments()
    main(*arguments)
