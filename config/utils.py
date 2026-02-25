import json
import pandas as pd
import requests


def load_companies():
    """Read the company config file"""
    with open("companies.json") as f:
        data = json.load(f)
        companies = data["companies"]
        return companies


def load_holidays():
    """Read the holiday dates from config file"""
    with open("holidays.json") as f:
        data = json.load(f)
        holidays = data["holidays"]
        return holidays


def get_eur_exchange_rate_nbp(start: pd.Timestamp, end: pd.Timestamp):
    """
    Get exchange rates from eur to pln (nbp, table A)
    from start date to end date.

    :param start: Beginning of a period to get rates
    :type start: pd.Timestamp
    :param end: End of a period to get rates
    :type end: pd.Timestamp
    """
    url = "http://api.nbp.pl/api/exchangerates/rates/a/eur/{}/{}/".format(
        start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
    )
    response = requests.get(url)

    if response.status_code == 200:
        data = response.json()
        rates = {}
        for r in data['rates']:
            rates[r["effectiveDate"]] = r["mid"]
        return rates
    else:
        return {}
