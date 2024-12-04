import datetime
from collections import defaultdict
import pandas as pd

def workingDays(startDate, endDate):
    working_dates = []
    working_dates_with_weekdays = []
    week_dates = defaultdict(list)
    month_dates = defaultdict(list)
    current_date = startDate
    
    while current_date <= endDate:
        if current_date.weekday() < 5:  # Check if it's a weekday
            formatted_date = current_date.strftime('%Y-%m-%d')
            working_dates.append(formatted_date)
            working_dates_with_weekdays.append((formatted_date, current_date.weekday()))

            week_number = current_date.isocalendar()[1]
            month_number = current_date.month
            
            week_dates[week_number].append(formatted_date)
            month_dates[month_number].append(formatted_date)
        
        current_date += datetime.timedelta(days=1)
    
    return {
        'workingDates': working_dates,
        'workingDatesWithWeekdays': working_dates_with_weekdays,
        'workingWeeks': dict(week_dates),
        'workingMonths': dict(month_dates)
    }

def get_month_number(month_name):
    return {
        'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
        'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
    }.get(month_name, 0)

def get_month_name(month_number):
    month_names = {
        1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June',
        7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'
    }
    return month_names.get(month_number, "")

def load_data(file_name):
    xls = pd.ExcelFile(file_name)
    assessors = {}
    candidate_goal = {}
    program_capacities = {}
    office_unavailabilities = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name)

        if sheet_name != 'Extra':
            assessor_data = {}
            for _, row in df.iterrows():
                key, value = row.iloc[0], row.iloc[1]
                if key.startswith('Capacity - '):
                    month = key.split('- ')[1]
                    month_number = get_month_number(month)
                    assessor_data.setdefault('Capacity', {})[month_number] = int(value)
                    assessor_data.setdefault('Capacity', {})[month_number] = int(value)

                elif key == 'Activities':
                    assessor_data[key] = value.split(', ') if value else []
                elif key in ['DATA', 'HR', 'weeklyUnavailability', 'Unavailability']:
                    if value in ['TRUE', 'FALSE']:
                        assessor_data[key] = value == 'TRUE'
                    elif pd.isna(value):
                        assessor_data[key] = []
                    else:
                        assessor_data[key] = [int(v) if v.isdigit() else v for v in value.split(', ')] if isinstance(value, str) else [value]
                elif key == 'programs':
                    assessor_data[key] = value.split(', ') if value else []
                elif key == 'assessmentAvailability':
                    if pd.notna(value):
                        assessor_data[key] = value.split(', ') if value else []
                    else:
                        print(f"No assessment availability found for {sheet_name}")
                elif key == 'caseAvailability':
                    if pd.notna(value):
                        assessor_data[key] = value.split(', ') if value else []   
                    else:
                        print(f"No case availability found for {sheet_name}")
                else:
                    assessor_data[key] = value
            assessors[sheet_name] = assessor_data
        else:
            for _, row in df.iterrows():
                key, value = row['Key'], row['Value']
                if pd.isna(key):
                    continue
                if 'Candidate Goal - ' in key:
                    month = key.split('- ')[-1]
                    month_number = get_month_number(month)
                    program_capacities[month] = {
                        'MCP&DATA': int(row['MCP&DATA']),
                        'AM IT': int(row['AM IT']),
                        'Buildwise': int(row['Buildwise']),
                        'Scrum Master': int(row['Scrum Master']),
                        'Pluxee': int(row['Pluxee']),
                        'Curious': int(row['Curious']),
                        ### Future programs
                        'Program1': int(row['Program1']),
                        'Program2': int(row['Program2']),
                        'Program3': int(row['Program3']),
                        'Program4': int(row['Program4']),
                        'Program5': int(row['Program5']),
                        'Program6': int(row['Program6']),
                        'Program7': int(row['Program7']),
                        'Program8': int(row['Program8']),
                        }
                elif key in ['Public Holidays', 'Office Events']:
                    if pd.notna(value):  # Check dates present
                        dates = value.split(', ')
                        office_unavailabilities.extend(dates)
           
    return assessors, program_capacities, office_unavailabilities