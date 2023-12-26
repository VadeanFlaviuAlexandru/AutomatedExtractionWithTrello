from trello import TrelloClient

from utils.config import API_KEY, API_SECRET, API_TOKEN
from utils.constants import MOBILE, BACKEND, FRONTEND, BUGS, TECHNICAL_REVIEW, TESTING, DOING, RETURNED_FROM_TESTING, \
    BLOCKED


def connect():
    trello_client = TrelloClient(
        api_key=API_KEY,
        api_secret=API_SECRET,
        token=API_TOKEN
    )

    return null


def extract_data():
    return null
