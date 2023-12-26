from data.data_processing import split_ticket_title
from sheets.text_config import textFormatConfig, alignBottomConfig
from utils.constants import ESTIMATIONS, PLURAL_ESTIMATIONS, TABLE_HEADERS
from utils.utils import insert_vowel_after_first_word, minutes_to_hours


def lists(workbook, boardCards: dict, titleFormat):
    worksheet = workbook.add_worksheet(name=f'Lists')

    textFormat = workbook.add_format(
        {**textFormatConfig, **alignBottomConfig, 'text_wrap': True, 'bg_color': '#EFEFEF'})

    totalsCellFormat = workbook.add_format(
        {**textFormatConfig, **alignBottomConfig, 'text_wrap': True, 'bg_color': '#666666'})

    totalRowFormat = workbook.add_format(
        {**textFormatConfig, **alignBottomConfig, 'text_wrap': True, 'bg_color': '#9B9B9B'})

    estimatedRowFormat = workbook.add_format(
        {**textFormatConfig, **alignBottomConfig, 'text_wrap': True, 'bg_color': '#BCBCBC'})

    unestimatedRowFormat = workbook.add_format(
        {**textFormatConfig, **alignBottomConfig, 'text_wrap': True, 'bg_color': '#ACACAC'})

    worksheet.set_column(first_col=0, last_col=0, width=35)
    worksheet.set_column(first_col=1, last_col=4, width=25)
    worksheet.set_column(first_col=5, last_col=5, width=35)
    worksheet.freeze_panes(1, 0)

    worksheet.write('A1', TABLE_HEADERS['listName'], titleFormat)
    worksheet.write('B1', TABLE_HEADERS['estimationType'], titleFormat)
    worksheet.write('C1', TABLE_HEADERS['hoursNumberXhYm'], titleFormat)
    worksheet.write('D1', TABLE_HEADERS['hoursNumberXZ'], titleFormat)
    worksheet.write('E1', TABLE_HEADERS['ticketsNr'], titleFormat)

    ticketsCounter = {}
    ticketsTotal = {}

    ticketsTotal.setdefault('toate', {'minutes': 0, 'ticketsNr': 0})
    ticketsTotal.setdefault(ESTIMATIONS['estimated'], {'minutes': 0, 'ticketsNr': 0})
    ticketsTotal.setdefault(ESTIMATIONS['unestimated'], {'minutes': 0, 'ticketsNr': 0})
    ticketsTotal.setdefault(ESTIMATIONS['invaluable'], {'minutes': 0, 'ticketsNr': 0})

    for key, listOfCards in boardCards.items():
        ticketsCounter[key] = {}

        ticketsCounter[key].setdefault('total', {'minutes': 0, 'ticketsNr': 0})
        ticketsCounter[key].setdefault(PLURAL_ESTIMATIONS['estimated'], {'minutes': 0, 'ticketsNr': 0})
        ticketsCounter[key].setdefault(PLURAL_ESTIMATIONS['unestimated'], {'minutes': 0, 'ticketsNr': 0})
        ticketsCounter[key].setdefault(PLURAL_ESTIMATIONS['unestimatedI'], {'minutes': 0, 'ticketsNr': 0})
        ticketsCounter[key].setdefault(PLURAL_ESTIMATIONS['unestimatedCT'], {'minutes': 0, 'ticketsNr': 0})
        ticketsCounter[key].setdefault(PLURAL_ESTIMATIONS['unestimatedICT'], {'minutes': 0, 'ticketsNr': 0})
        ticketsCounter[key].setdefault(PLURAL_ESTIMATIONS['invaluable'], {'minutes': 0, 'ticketsNr': 0})

        for card in listOfCards:
            implementationEstimate, reviewAndTestingEstimate, implementationOverrun, reviewAndTestingOverrun, _, \
                _, estimationType, _ = split_ticket_title(card.name)

            pluralEstimation = insert_vowel_after_first_word(estimationType, 'e')

            if estimationType == ESTIMATIONS['estimated']:
                ticketsCounter[key][pluralEstimation]['minutes'] += (implementationEstimate or 0) + (
                        reviewAndTestingEstimate or 0) + (implementationOverrun or 0) + \
                                                                    (reviewAndTestingOverrun or 0)
                ticketsCounter[key][pluralEstimation]['ticketsNr'] += 1

            if estimationType == ESTIMATIONS['unestimatedI']:
                ticketsCounter[key][pluralEstimation]['minutes'] += (reviewAndTestingEstimate or 0) + (
                        reviewAndTestingOverrun or 0)
                ticketsCounter[key][pluralEstimation]['ticketsNr'] += 1

            if estimationType == ESTIMATIONS['unestimatedCT']:
                ticketsCounter[key][pluralEstimation]['minutes'] += (implementationEstimate or 0) + (
                        implementationOverrun or 0)
                ticketsCounter[key][pluralEstimation]['ticketsNr'] += 1

            if estimationType == ESTIMATIONS['invaluable']:
                ticketsCounter[key][pluralEstimation]['ticketsNr'] += 1

            if estimationType == ESTIMATIONS['unestimatedICT']:
                ticketsCounter[key][PLURAL_ESTIMATIONS['unestimated']]['ticketsNr'] += 1
                ticketsCounter[key][pluralEstimation]['ticketsNr'] += 1

        estimatedTime = ticketsCounter[key].get(PLURAL_ESTIMATIONS['estimated'], None).get('minutes')

        estimatedTickets = ticketsCounter[key].get(PLURAL_ESTIMATIONS['estimated'], None).get('ticketsNr')

        unvaluableTickets = ticketsCounter[key].get(PLURAL_ESTIMATIONS['invaluable'], None).get('ticketsNr')

        ticketsCounter[key][PLURAL_ESTIMATIONS['unestimated']]['minutes'] = \
            ticketsCounter[key][PLURAL_ESTIMATIONS['unestimatedCT']]['minutes'] + \
            ticketsCounter[key][PLURAL_ESTIMATIONS['unestimatedI']]['minutes']

        ticketsCounter[key][PLURAL_ESTIMATIONS['unestimated']]['ticketsNr'] = \
            ticketsCounter[key][PLURAL_ESTIMATIONS['unestimatedCT']]['ticketsNr'] + \
            ticketsCounter[key][PLURAL_ESTIMATIONS['unestimatedI']]['ticketsNr'] + \
            ticketsCounter[key][PLURAL_ESTIMATIONS['unestimatedICT']]['ticketsNr']

        unestimatedTime = ticketsCounter[key].get(PLURAL_ESTIMATIONS['unestimated'], None).get('minutes')

        unestimatedTickets = ticketsCounter[key].get(PLURAL_ESTIMATIONS['unestimated'], None).get('ticketsNr')

        ticketsCounter[key]['total']['minutes'] = estimatedTime + unestimatedTime

        ticketsCounter[key]['total']['ticketsNr'] = estimatedTickets + unestimatedTickets + unvaluableTickets

        ticketsTotal[ESTIMATIONS['estimated']]['minutes'] += estimatedTime

        ticketsTotal[ESTIMATIONS['unestimated']]['minutes'] += unestimatedTime

        ticketsTotal['toate']['minutes'] += estimatedTime + unestimatedTime

        ticketsTotal[ESTIMATIONS['estimated']]['ticketsNr'] += estimatedTickets

        ticketsTotal[ESTIMATIONS['unestimated']]['ticketsNr'] += unestimatedTickets

        ticketsTotal['toate']['ticketsNr'] += estimatedTickets + unestimatedTickets + unvaluableTickets

        ticketsTotal[ESTIMATIONS['invaluable']]['ticketsNr'] += unvaluableTickets

    row = 1

    for key, value in ticketsTotal.items():
        worksheet.write_string(row=row, col=0, string='toate', cell_format=totalsCellFormat)
        worksheet.write_string(row=row, col=1, string=key or '', cell_format=totalsCellFormat)
        worksheet.write_string(row=row, col=2, string=minutes_to_hours(value['minutes']),
                               cell_format=totalsCellFormat)
        worksheet.write_string(row=row, col=3, string=str(round(value['minutes'] / 60, 2)),
                               cell_format=totalsCellFormat)
        worksheet.write_string(row=row, col=4, string=str(value['ticketsNr']),
                               cell_format=totalsCellFormat)

        row += 1

    row += 1

    for key, lists in ticketsCounter.items():

        for estType, hours in lists.items():
            if row % 8 == 6:
                cellFormat = totalRowFormat
            elif row % 8 == 0 or row % 8 == 7:
                cellFormat = estimatedRowFormat
            elif row % 8 == 4:
                cellFormat = unestimatedRowFormat
            else:
                cellFormat = textFormat

            worksheet.write_string(row=row, col=0, string=key, cell_format=cellFormat)
            worksheet.write_string(row=row, col=1, string=estType or '', cell_format=cellFormat)
            worksheet.write_string(row=row, col=2, string=minutes_to_hours(hours['minutes']),
                                   cell_format=cellFormat)
            worksheet.write_string(row=row, col=3, string=str(round(hours['minutes'] / 60, 2)),
                                   cell_format=cellFormat)
            worksheet.write_string(row=row, col=4, string=str(hours['ticketsNr']),
                                   cell_format=cellFormat)
            row += 1

        row += 1

    return worksheet
