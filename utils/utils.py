import re
from datetime import datetime

from utils.constants import INACTIVE_PROJECTS, INACTIVE_LABELS, NON_PROJECT_LABELS


def get_valid_labels(labels: list, removedLabels):
    matchedLabels = list(
        filter(lambda label: label.name.lower() not in removedLabels,
               labels))

    return matchedLabels


def get_all_valid_labels(boardCards):
    valid_project_labels = set()

    for key, listOfCards in boardCards.items():
        for card in listOfCards:
            validLabels = [label.name for label in
                           get_valid_labels(card.labels, INACTIVE_PROJECTS + INACTIVE_LABELS + NON_PROJECT_LABELS)]
            valid_project_labels.update(validLabels)

    return valid_project_labels


def days_since_card_creation_date(targetDate: datetime):
    currentDate = datetime.now()
    timeDifference = currentDate - targetDate

    numberOfDays = timeDifference.days

    return numberOfDays


def contains_only_one_from_list(string, stringList):
    count = 0
    for item in stringList:
        if item.lower() == string.lower():
            count += 1
            if count > 1:
                return False
    return count == 1


def replace_except_whitelist(inputString, allowedChars="+-:hHmM1234567890. ", exceptList=None):
    if exceptList is None:
        exceptList = []

    for substring in exceptList:
        inputString = re.sub(rf"(?<!\w){re.escape(substring)}(?!\w)", '', inputString)

    pattern = rf"(?P<keep>min)|[^{re.escape(allowedChars)}]"
    resultString = re.sub(pattern, lambda x: x.group("keep") or '', inputString)

    return resultString


def ends_with_substring_with_space(string, substrings):
    for substring in substrings:
        if string.endswith(substring):
            index = string.rfind(substring)
            if index == 0:
                continue

            beforeSubstring = string[index - 1]
            if beforeSubstring.isspace():
                return True
    return False


def remove_substrings(string, substrings):
    sortedSubstrings = sorted(substrings, key=len, reverse=True)
    resultString = string
    for substring in sortedSubstrings:
        resultString = resultString.replace(substring, '')
    return resultString


def contains_numbers(inputString):
    for char in inputString:
        if char.isnumeric():
            return True
    return False


def time_string_to_minutes(timeStr):
    totalMinutes = 0

    pattern = r'(\d+(\.\d+)?)(h|m|min|:h|:m|:)?'
    matches = re.findall(pattern, timeStr)

    for i, (number, _, unit) in enumerate(matches):
        number = float(number) if '.' in number else int(number)

        if i > 0 and re.match(r'\d', unit) and re.match(r'\d', matches[i - 1][2]):
            totalMinutes += (number * 60) + time_string_to_minutes(matches[i - 1][0] + 'h')

        elif unit is None or unit == 'h':
            totalMinutes += number * 60
        elif unit == 'm' or unit == ':m':
            totalMinutes += number
        elif unit == ':h':
            totalMinutes += number * 60
        elif unit == ':':
            totalMinutes += number * 60

    return totalMinutes


def evaluate_time_expression(expression):
    expression = expression.replace('min', 'm')
    parts = re.split(r'\s*([-+])\s*', expression)

    estimation = 0
    extraMinutes = 0

    if len(parts) > 0:
        myParts = parts[0]
        if is_valid_integer(parts[0]):
            myParts = parts[0] + 'h'
        if ':' in parts[0]:
            myParts = convert_time_format(parts[0])

        estimation = time_string_to_minutes(myParts)

        for i in range(0, len(parts), 2):
            myParts = parts[i]
            if is_valid_integer(parts[i]):
                myParts = parts[i] + 'h'
            if ':' in parts[i]:
                myParts = convert_time_format(parts[i])
            sign = 1 if parts[i - 1] == '+' else -1
            if len(parts) > 1 and i > 1:
                extraMinutes += sign * time_string_to_minutes(myParts)

    return int(estimation), int(extraMinutes)


def is_valid_integer(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


def convert_time_format(inputStr):
    hours, minutes = inputStr.split(':')
    hours = int(hours)
    minutes = int(minutes.rstrip('hm'))

    formattedTime = f"{hours}h"
    if minutes > 0:
        formattedTime += f"{minutes}m"

    return formattedTime


def are_square_brackets_closed(inputStr):
    stack = []

    for char in inputStr:
        if char == "[":
            stack.append(char)
        elif char == "]":
            if not stack or stack.pop() != "[":
                return False

    return len(stack) == 0


def count_empty_brackets(string):
    emptyBrackets = re.findall(r'\[\s*\]', string)
    return len(emptyBrackets)


def is_ordered(lst, descending=False):
    if not lst:
        return True

    for i in range(len(lst) - 1):
        if (descending and lst[i] < lst[i + 1]) or (not descending and lst[i] > lst[i + 1]):
            return False

    return True


def remove_after_last_slash(link):
    lastSlashIndex = link.rfind('/')
    if lastSlashIndex != -1:
        return link[:lastSlashIndex + 1]
    return link


def insert_vowel_after_first_word(input_string, vowel):
    words = input_string.split()
    if len(words) >= 2:
        return words[0] + vowel + ' ' + ' '.join(words[1:])
    else:
        return input_string + vowel


def minutes_to_hours(minutes):
    hours = minutes // 60
    minutes_remainder = minutes % 60

    hours = hours > 0 and f"{hours}h" or ""

    return minutes != 0 and f"{hours}{minutes_remainder:02d}m" or "0.00"


def add_missing_keys(dictionary, keys, default_value=None):
    for key in keys:
        if key not in dictionary:
            dictionary[key] = default_value