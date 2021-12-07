import openpyxl
from pathlib import Path
from itertools import combinations

from simpleai.search import (
    CspProblem,
    backtrack,
    HIGHEST_DEGREE_VARIABLE,
    MOST_CONSTRAINED_VARIABLE,
    LEAST_CONSTRAINING_VALUE,
)

preference_sheet_file = Path('data','Preferences 2022 CSW Staff.xlsx')
# print(preference_sheet_file)

workbook_obj = openpyxl.load_workbook(preference_sheet_file)
preference_sheet = workbook_obj.active

# Get the Staff names
staff = []
for row in preference_sheet.iter_rows(min_row=3, max_row=30, max_col=1):
    for cell in row:
        staff.append(cell.value)

# Get the full list of camp activities
activities = []
for col in preference_sheet.iter_cols(min_row=2, max_row=2, max_col=preference_sheet.max_column):
    for cell in col:
        if cell.value != None:
            activities.append(cell.value)

# Get preferences and creates a dictionary
preferences = {}
# CELLS_TO_IGNORE = {14,17,60,77,80,153}
CELLS_TO_IGNORE = {12,15,58,75,78,151}

for row_index, row in enumerate(preference_sheet.iter_rows(min_row=3, max_row=30, min_col=2, max_col=169)):
    temp_preference_list = []
    for cell_index, cell in enumerate(row):
        if cell_index not in CELLS_TO_IGNORE:
            if cell.value != None:
                # print(cell.value)
                if cell.value.casefold() == "lead":
                    temp_preference_list.append(2)
                elif cell.value.casefold() == "assist":
                    temp_preference_list.append(1)
                else:
                    temp_preference_list.append(0)
            else: # If there is no preference set, a default 1 (ASSIST) will be set
                temp_preference_list.append(1)
    preferences[staff[row_index]] = temp_preference_list




#print(preference_sheet["D2"].value)