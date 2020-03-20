#-------------------------------------------------------------------------------
# Name:        Yu-Gi-Oh! TCG Scraper
# Purpose:     Scrapes the Yugioh Prices TCG API for all cards and renders them
#              in a desired Excel format.
#
# Author:      Zak Turchansky
#
# Created:     15-06-2019
# Copyright:   (c) Zak Turchansky 2019
# Licence:     LICENSE.txt
#-------------------------------------------------------------------------------
"""
Scrapes the Yugioh Prices TCG API for all cards and renders them in an Excel
spreadsheet.
"""
import urllib.parse
from os import remove
from pathlib import Path
import requests
from openpyxl import Workbook
from openpyxl.styles import Font

API_BASE_URL = 'http://yugiohprices.com/api/'

def main():
    """Add all Yu-Gi-Oh! TCG card information from every set to an Excel file"""
    card_fields = ["name",
                   "text",
                   "card_type",
                   "type",
                   "family",
                   "atk",
                   "def",
                   "level",
                   "property"]
    set_fields = ["name",
                  "print_tag",
                  "rarity"]

    workbook = create_workbook(card_fields, set_fields)

    sets = get_sets()

    for card_set in sets:
        add_cards_in_set_to_workbook(card_set,
                                     workbook,
                                     card_fields,
                                     set_fields)

    workbook.save("ygo_output.xlsx")

def get_sets():
    """Return a list of all Yu-Gi-Oh TCG set names."""
    sets_url = API_BASE_URL + "card_sets"
    print("Fetching set names...")
    sets = requests.get(sets_url).json()
    return sets

def add_cards_in_set_to_workbook(set_name, workbook, card_fields, set_fields):
    """Add a row to the workbook for each card in the given set."""
    print("Fetching cards from " + set_name)
    set_url = API_BASE_URL + "set_data/" + set_name
    card_url = API_BASE_URL + "card_data/"
    cards_response = requests.get(set_url).json()

    for card_setinfo in cards_response['data']['cards']:
        card_row = []
        url = card_url + urllib.parse.quote_plus(card_setinfo['name'])
        card_info_response = requests.get(url).json()
        card_info = card_info_response['data']

        for card_field in card_fields:
            card_row.append(card_info[card_field])
        for set_field in set_fields:
            card_row.append(card_setinfo['numbers'][0][set_field])
        workbook.active.append(card_row)

def create_workbook(card_fields, set_fields):
    """Create an empty excel workbook with a header row and return it."""
    # Update the set names for the column headers since they are both just
    # 'name' when we fetch them from the API
    header_fields = card_fields + set_fields
    header_fields[0] = 'card_name'
    header_fields[len(card_fields)] = 'set_name'
    if Path("ygo_output.xlsx").is_file():
        remove("ygo_output.xlsx")
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(header_fields)
    for cell in worksheet["1:1"]:
        cell.font = Font(bold=True)
    return workbook

if __name__ == '__main__':
    main()
