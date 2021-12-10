import os
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

HEAD_COUNSELORS = {'Brodi','Ceci','Ignacio','Jocelyn','John','Lily','Max','Micky','Noor'}
COUNSELORS = {'Alicja','Brent','Daniel','Dom','Gabriel','Ildar','Juliana','Murray','Nelson','Simone','Sparta','Sydy'}
INTERNS = {'Audrey','Ben','Jack','Jordan','Nathan','Sophie','Van'}
ARCHERY_CERT = {'Ignacio','Jocelyn','Dom','Gabriel','Sophie','Jack','Brent','Daniel'}
LIFEGUARD_CERT = {'Brodi','Ceci','Dom','Ildar','Sparta','Nathan','Van','Ben'}
# CELLS_TO_IGNORE = {14,17,60,77,80,153}
CELLS_TO_IGNORE = {12,15,58,75,78,151} # Original -2

def spreadsheet_tasks():
    # Get preferences spreadsheet
    preference_sheet_file = Path(os.path.dirname(os.path.abspath(__file__)),'data','Preferences 2022 CSW Staff.xlsx')
    preference_workbook_obj = openpyxl.load_workbook(preference_sheet_file)
    preference_sheet = preference_workbook_obj.active

    #Get tomorrow's schedule
    schedule_sheet_file = Path(os.path.dirname(os.path.abspath(__file__)),'data','Schedule_1P.xlsx')
    schedule_workbook_obj = openpyxl.load_workbook(schedule_sheet_file)
    schedule_sheet = schedule_workbook_obj.active    

    # Get list of staff
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
    activities.append('Free Period')

    # Get the list of activities from tomorrow's schedule
    schedule_activities = []
    for row_index, row in enumerate(schedule_sheet.iter_rows(max_row=schedule_sheet.max_row, max_col=schedule_sheet.max_column)):
        for cell in row:
            if cell.value != None:
                schedule_activities.append(cell.value)
                # schedule_activities.append([row_index + 1,cell.value])
    schedule_activities.append('Free Period')

    # Get preferences and creates a dictionary
    preferences = {}
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
        # print(preferences)

    return staff, activities, preferences, schedule_activities

staff, activities, preferences, schedule_activities = spreadsheet_tasks()

# VARIABLES AND DOMAINS

problem_variables = staff
domains = {}
for variable in problem_variables:
    if variable in HEAD_COUNSELORS:
        domains[variable] = ['Free Period']
    else:
        domains[variable] = schedule_activities

# CONSTRAINTS

constraints = []

def lifeguards_certs_on_waterfront(variables, values):
    pass

def archery_certs_on_archery(variables, values):
    pass

def not_just_interns(variables, values):
    pass

def take_preferences_into_account(variables, values):
    pass

#print(preference_sheet["D2"].value)