from data.data_processing import split_ticket_title
from sheets.text_config import textFormatConfig, alignBottomConfig
from utils.constants import INACTIVE_PROJECTS, INACTIVE_LABELS, NON_PROJECT_LABELS, TABLE_HEADERS, ESTIMATIONS
from utils.utils import get_valid_labels, minutes_to_hours, get_all_valid_labels


def labels(workbook, boardCards: dict, titleFormat):
    worksheet = workbook.add_worksheet(name=f'Labels')

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

    worksheet.set_column(first_col=0, last_col=4, width=20)

    worksheet.freeze_panes(1, 0)

    worksheet.write('A1', TABLE_HEADERS['labelName'], titleFormat)
    worksheet.write('B1', TABLE_HEADERS['estimationType'], titleFormat)
    worksheet.write('C1', TABLE_HEADERS['hoursNumberXhYm'], titleFormat)
    worksheet.write('D1', TABLE_HEADERS['hoursNumberXZ'], titleFormat)
    worksheet.write('E1', TABLE_HEADERS['ticketsNr'], titleFormat)

    labels = {}
    ticketsTotal = {}

    ticketsTotal.setdefault('toate', {'minutes': 0, 'ticketsNr': 0})
    ticketsTotal.setdefault(ESTIMATIONS['estimated'], {'minutes': 0, 'ticketsNr': 0})
    ticketsTotal.setdefault(ESTIMATIONS['unestimated'], {'minutes': 0, 'ticketsNr': 0})
    ticketsTotal.setdefault(ESTIMATIONS['invaluable'], {'minutes': 0, 'ticketsNr': 0})

    result = get_all_valid_labels(boardCards)

    for label in result:
        for key, listOfCards in boardCards.items():
            for card in listOfCards:
                implementationEstimate, reviewAndTestingEstimate, implementationOverrun, reviewAndTestingOverrun, _, \
                    _, estimationType, _ = split_ticket_title(card.name)

                cardLabels = get_valid_labels(card.labels, INACTIVE_PROJECTS + INACTIVE_LABELS + NON_PROJECT_LABELS)

                cardLabels = list(map(lambda obj: obj.name, cardLabels))

                if label in cardLabels:
                    if (label, "total") not in labels:
                        labels[(label, "total")] = {'minutes': 0, 'ticketsNr': 0}
                        labels[(label, ESTIMATIONS['estimated'])] = {'minutes': 0, 'ticketsNr': 0}
                        labels[(label, ESTIMATIONS['unestimated'])] = {'minutes': 0, 'ticketsNr': 0}
                        labels[(label, ESTIMATIONS['unestimatedI'])] = {'minutes': 0, 'ticketsNr': 0}
                        labels[(label, ESTIMATIONS['unestimatedCT'])] = {'minutes': 0, 'ticketsNr': 0}
                        labels[(label, ESTIMATIONS['unestimatedICT'])] = {'minutes': 0, 'ticketsNr': 0}
                        labels[(label, ESTIMATIONS['invaluable'])] = {'minutes': 0, 'ticketsNr': 0}

                    if estimationType == ESTIMATIONS['estimated']:
                        labels[(label, estimationType)]['minutes'] += (implementationEstimate or 0) + (
                                reviewAndTestingEstimate or 0) + (implementationOverrun or 0) + \
                                                                      (reviewAndTestingOverrun or 0)
                        labels[(label, estimationType)]['ticketsNr'] += 1

                    if estimationType == ESTIMATIONS['unestimatedI']:
                        labels[(label, estimationType)]['minutes'] += (reviewAndTestingEstimate or 0) + (
                                reviewAndTestingOverrun or 0)
                        labels[(label, estimationType)]['ticketsNr'] += 1

                    if estimationType == ESTIMATIONS['unestimatedCT']:
                        labels[(label, estimationType)]['minutes'] += (implementationEstimate or 0) + (
                                implementationOverrun or 0)
                        labels[(label, estimationType)]['ticketsNr'] += 1

                    if estimationType == ESTIMATIONS['invaluable']:
                        labels[(label, estimationType)]['ticketsNr'] += 1

                    if estimationType == ESTIMATIONS['unestimatedICT']:
                        labels[(label, ESTIMATIONS['unestimated'])]['ticketsNr'] += 1
                        labels[(label, estimationType)]['ticketsNr'] += 1

        estimatedTime = labels.get((label, ESTIMATIONS['estimated']), None).get('minutes')

        estimatedTickets = labels.get((label, ESTIMATIONS['estimated']), None).get('ticketsNr')

        unvaluableTickets = labels.get((label, ESTIMATIONS['invaluable']), None).get('ticketsNr')

        labels[(label, ESTIMATIONS['unestimated'])]['minutes'] = \
            labels[(label, ESTIMATIONS['unestimatedCT'])]['minutes'] + \
            labels[(label, ESTIMATIONS['unestimatedI'])]['minutes']

        labels[(label, ESTIMATIONS['unestimated'])]['ticketsNr'] = \
            labels[(label, ESTIMATIONS['unestimatedCT'])]['ticketsNr'] + \
            labels[(label, ESTIMATIONS['unestimatedI'])]['ticketsNr'] + \
            labels[(label, ESTIMATIONS['unestimatedICT'])]['ticketsNr']

        unestimatedTime = labels.get((label, ESTIMATIONS['unestimated']), None).get('minutes')

        unestimatedTickets = labels.get((label, ESTIMATIONS['unestimated']), None).get('ticketsNr')

        labels[(label, 'total')]['minutes'] = estimatedTime + unestimatedTime

        labels[(label, 'total')]['ticketsNr'] = estimatedTickets + unestimatedTickets + unvaluableTickets

        ticketsTotal[ESTIMATIONS['estimated']]['minutes'] += estimatedTime

        ticketsTotal[ESTIMATIONS['unestimated']]['minutes'] += unestimatedTime

        ticketsTotal['toate']['minutes'] += estimatedTime + unestimatedTime

        ticketsTotal[ESTIMATIONS['estimated']]['ticketsNr'] += estimatedTickets

        ticketsTotal[ESTIMATIONS['unestimated']]['ticketsNr'] += unestimatedTickets

        ticketsTotal['toate']['ticketsNr'] += estimatedTickets + unestimatedTickets + unvaluableTickets

        ticketsTotal[ESTIMATIONS['invaluable']]['ticketsNr'] += unvaluableTickets

    sorted_labels = sorted(labels.items(), key=lambda item: item[0][0])
    sorted_labels_dict = {label_key: label_value for label_key, label_value in sorted_labels}

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

    for labelEst, value in sorted_labels_dict.items():
        if row % 8 == 6:
            cellFormat = totalRowFormat
        elif row % 8 == 0 or row % 8 == 7:
            cellFormat = estimatedRowFormat
        elif row % 8 == 4:
            cellFormat = unestimatedRowFormat
        elif row % 8 == 5:
            row += 1
        else:
            cellFormat = textFormat

        label = labelEst[0] or ''
        estimation = labelEst[1] or ''
        worksheet.write_string(row=row, col=0, string=label, cell_format=cellFormat)
        worksheet.write_string(row=row, col=1, string=estimation or '', cell_format=cellFormat)
        worksheet.write_string(row=row, col=2, string=minutes_to_hours(value['minutes']), cell_format=cellFormat)
        worksheet.write_string(row=row, col=3, string=str(round(value['minutes'] / 60, 2)),
                               cell_format=cellFormat)
        worksheet.write_string(row=row, col=4, string=str(value['ticketsNr']),
                               cell_format=cellFormat)

        row += 1

    return worksheet
