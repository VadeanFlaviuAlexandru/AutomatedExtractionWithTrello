import os
import shutil
from datetime import datetime

import xlsxwriter

from data.data_extraction import extract_data
from sheets.labels import labels
from sheets.lists import lists
from sheets.text_config import textFormatConfig
from sheets.ticket import ticket


def generate_excel():
    board_cards = extract_data()

    excel_filename = "Trello_Daily_{}.xlsx".format(datetime.now().strftime('%Y-%d-%m_%Hh%Mm%Ss'))
    workbook = xlsxwriter.Workbook(excel_filename)
    title_format = workbook.add_format({'bold': 1, **textFormatConfig, 'bottom': 2})

    ticket(workbook, board_cards, title_format)
    lists(workbook, board_cards, title_format)
    labels(workbook, board_cards, title_format)
    workbook.close()

    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")

    shutil.move(excel_filename, os.path.join(downloads_path, excel_filename))
