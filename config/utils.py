import json


def load_companies():
    """Read the company config file"""
    with open("companies.json") as f:
        data = json.load(f)
        companies = data["companies"]
        return companies
