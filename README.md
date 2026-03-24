# MTG Collection Matcher

A Python tool that compares a Magic: The Gathering decklist against an Excel-based card collection.

## What it does

- Reads a decklist from a text file (format: `1 Card Name`)
- Scans an Excel file containing a card collection
- Matches cards by name, including variants like:
  - Swamp
  - Swamp (Foil)
  - Swamp (Extended)
- Filters out cards that are marked as unavailable
- Outputs:
  - cards available in the collection (with quantities)
  - cards that are missing or unavailable

## How it works

- Uses Python and the `openpyxl` library to read Excel data
- Normalizes card names so different versions match correctly
- Stores collection data in a dictionary for fast lookup
- Uses a status column in the Excel sheet to determine availability

## How to run

- Install the required dependency:

- pip install -r requirements.txt

- (If pip does not work, use pip3)

- Then run the program:

- python3 match_cards.py
