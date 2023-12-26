from data.data_processing import split_ticket_title
from sheets.text_config import textFormatConfig, alignBottomConfig
from utils.constants import INACTIVE_PROJECTS, INACTIVE_LABELS, TABLE_HEADERS
from utils.utils import get_valid_labels, days_since_card_creation_date, remove_after_last_slash


def ticket(workbook, boardCards: dict, titleFormat):
    worksheet = workbook.add_worksheet(name=f'Ticket')

    textFormat = workbook.add_format({**textFormatConfig, **alignBottomConfig, 'text_wrap': True})

    redTextFormat = workbook.add_format({**textFormatConfig, **alignBottomConfig, 'text_wrap': True, 'color': 'red'})

    greenTextFormat = workbook.add_format(
        {**textFormatConfig, **alignBottomConfig, 'text_wrap': True, 'color': 'green'})

    urlTextFormat = workbook.add_format({**textFormatConfig, 'color': 'blue'})

    urlErrorFormat = workbook.add_format({**textFormatConfig, 'color': 'red'})

    worksheet.set_column(first_col=0, last_col=0, width=30)
    worksheet.set_column(first_col=1, last_col=1, width=5)
    worksheet.set_column(first_col=2, last_col=2, width=10)
    worksheet.set_column(first_col=3, last_col=6, width=20)
    worksheet.set_column(first_col=7, last_col=7, width=25)
    worksheet.set_column(first_col=8, last_col=8, width=15)
    worksheet.set_column(first_col=9, last_col=9, width=10)
    worksheet.set_column(first_col=10, last_col=10, width=30)
    worksheet.freeze_panes(1, 0)

    worksheet.write('A1', TABLE_HEADERS['listName'], titleFormat)
    worksheet.write('B1', TABLE_HEADERS['size'], titleFormat)
    worksheet.write('C1', TABLE_HEADERS['estimationType'], titleFormat)
    worksheet.write('D1', TABLE_HEADERS['estimationI'], titleFormat)
    worksheet.write('E1', TABLE_HEADERS['overrunI'], titleFormat)
    worksheet.write('F1', TABLE_HEADERS['estimationCT'], titleFormat)
    worksheet.write('G1', TABLE_HEADERS['overrunCT'], titleFormat)
    worksheet.write('H1', TABLE_HEADERS['ticketTitle'], titleFormat)
    worksheet.write('I1', TABLE_HEADERS['label'], titleFormat)
    worksheet.write('J1', TABLE_HEADERS['daysSinceCreation'], titleFormat)
    worksheet.write('K1', TABLE_HEADERS['ticketURL'], titleFormat)

    row = 2
    for key, listOfCards in boardCards.items():
        for card in listOfCards:
            implementationEstimate, reviewAndTestingEstimate, implementationOverrun, reviewAndTestingOverrun, size, \
                title, estimationType, error = split_ticket_title(card.name)

            estimationCellFormat = redTextFormat if implementationOverrun and implementationOverrun > 0 else (
                greenTextFormat if implementationOverrun is not None and implementationOverrun != 0 else textFormat)

            daysSinceCreated = days_since_card_creation_date(card.card_created_date)

            daysCellFormat = redTextFormat if daysSinceCreated >= 7 else textFormat

            urlFormat = urlErrorFormat if error else urlTextFormat

            validLabels = ', '.join(
                [label.name for label in get_valid_labels(card.labels, INACTIVE_PROJECTS + INACTIVE_LABELS)])
            worksheet.write(f"A{row}", key, textFormat)
            worksheet.write(f"B{row}", size or '', textFormat)
            worksheet.write(f"C{row}", estimationType or "", textFormat)
            worksheet.write(f"D{row}", implementationEstimate and round(implementationEstimate / 60, 2) or 0,
                            implementationEstimate and implementationEstimate > 240 and redTextFormat or textFormat)
            worksheet.write(f"E{row}", implementationOverrun and round(implementationOverrun / 60, 2) or 0,
                            estimationCellFormat)
            worksheet.write(f"F{row}", reviewAndTestingEstimate and round(reviewAndTestingEstimate / 60, 2) or 0,
                            textFormat)
            worksheet.write(f"G{row}", reviewAndTestingOverrun and round(reviewAndTestingOverrun / 60, 2) or 0,
                            textFormat)
            worksheet.write(f"H{row}", title or '', textFormat)
            worksheet.write(f"I{row}", validLabels or '', textFormat)
            worksheet.write(f"J{row}", daysSinceCreated, daysCellFormat)
            worksheet.write_url(f"K{row}", card.url and remove_after_last_slash(card.url) or '', cell_format=urlFormat)

            row += 1

    return worksheet
