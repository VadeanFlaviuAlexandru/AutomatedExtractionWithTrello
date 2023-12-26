from utils.constants import CODE_REVIEW_TESTING_LABELS, SIZE_LABELS, ESTIMATIONS
from utils.utils import count_empty_brackets, are_square_brackets_closed
from utils.utils import replace_except_whitelist, \
    ends_with_substring_with_space, remove_substrings, contains_numbers, evaluate_time_expression, \
    contains_only_one_from_list, is_valid_integer, is_ordered, time_string_to_minutes


def split_ticket_title(ticketName):
    implementationEstimate = None
    reviewAndTestingEstimate = None
    implementationOverrun = None
    reviewAndTestingOverrun = None
    estimationType = None
    size = None
    title = None
    error = False
    currentCellsPosition = []
    ticketNameSplitted = ticketName.split("]")
    nameSplittedFields = []

    for str in ticketNameSplitted:
        nameSplittedFields.extend(str.split('['))

    for index, field in enumerate(nameSplittedFields):
        if 'time_est' in field:
            implementationEstimate = 0
            currentCellsPosition.append(0)
        elif is_valid_integer(field.replace(' ', '')):
            implementationEstimate = time_string_to_minutes(field.replace(' ', '') + 'h')
            error = True
        elif contains_only_one_from_list(field.replace(" ", ""), SIZE_LABELS):
            size = field.replace(" ", "")
            currentCellsPosition.append(2)
        elif replace_except_whitelist(field,
                                      "CcTt+-:hHmM1234567890 ") == field and ends_with_substring_with_space(
            field,
            CODE_REVIEW_TESTING_LABELS):
            if replace_except_whitelist(field,
                                        "CcTt+ ") == field:
                reviewAndTestingEstimate = 0
            else:
                newField = remove_substrings(field, CODE_REVIEW_TESTING_LABELS)
                reviewAndTestingEstimate, reviewAndTestingOverrun = evaluate_time_expression(
                    newField)
            currentCellsPosition.append(1)
        elif replace_except_whitelist(field,
                                      "+-:hHmM1234567890. ", ['min']) == field and contains_numbers(field):
            implementationEstimate, implementationOverrun = evaluate_time_expression(field)
            currentCellsPosition.append(0)

    if not are_square_brackets_closed(ticketName) or not is_ordered(currentCellsPosition):
        error = True

    title = ticketNameSplitted and ticketNameSplitted[-1].strip()

    if implementationEstimate and implementationEstimate > 0:
        if reviewAndTestingEstimate == 0:
            estimationType = ESTIMATIONS['unestimatedCT']
        elif reviewAndTestingEstimate and reviewAndTestingEstimate > 0 or reviewAndTestingEstimate is None:
            estimationType = ESTIMATIONS['estimated']
    elif implementationEstimate == 0:
        if reviewAndTestingEstimate and reviewAndTestingEstimate > 0 or reviewAndTestingEstimate is None:
            estimationType = ESTIMATIONS['unestimatedI']
        else:
            estimationType = ESTIMATIONS['unestimatedICT']
    elif implementationEstimate is None:
        if reviewAndTestingEstimate == 0:
            estimationType = ESTIMATIONS['unestimatedICT']
            error = True
        elif reviewAndTestingEstimate and reviewAndTestingEstimate > 0:
            estimationType = ESTIMATIONS['unestimatedI']
            error = True
        else:
            estimationType = ESTIMATIONS['invaluable']

    if count_empty_brackets(ticketName) == 1:
        if implementationEstimate and implementationEstimate > 0 and not (
                reviewAndTestingEstimate and reviewAndTestingEstimate > 0):
            estimationType = 'neestimat C+T'
        elif reviewAndTestingEstimate and reviewAndTestingEstimate > 0 and not (
                implementationEstimate and implementationEstimate > 0):
            estimationType = 'neestimat I'
    elif count_empty_brackets(ticketName) == 2:
        if implementationEstimate and implementationEstimate > 0 and reviewAndTestingEstimate is None:
            estimationType = 'neestimat C+T'
        elif reviewAndTestingEstimate and reviewAndTestingEstimate > 0 and implementationEstimate is None:
            estimationType = 'neestimat I'
        else:
            estimationType = 'neestimat I & C+T'

    return implementationEstimate, reviewAndTestingEstimate, implementationOverrun, reviewAndTestingOverrun, size, title, estimationType, error
