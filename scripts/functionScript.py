import pandas as pd
from ortools.sat.python import cp_model
import os
import sys
from functions import *
from datetime import datetime
from ics import Calendar, Event, Attendee
import pytz
from tkinter import filedialog, END
import json
        
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))
    
def log_message(message, output_text):
    output_text.configure(state='normal')  # Enable editing
    output_text.insert(END, f"{message}\n")  # Insert message at the end
    output_text.configure(state='disabled')  # Disable editing after update
    output_text.yview(END)  # Scroll to the latest message
    output_text.update_idletasks()  # Force the GUI to update immediately

def makeICS(schedule, startdate, enddate, email_file='emails.json'):
    # Load emails from the JSON file
    email_file = os.path.join(base_path, 'resources', 'emails.json')
    with open(email_file, 'r') as f:
        emails = json.load(f)

    timezone = pytz.timezone("Europe/Brussels")
    
    # Create a dictionary to hold events by program
    programs = {}

    # Loop through the schedule to create events
    for index, row in schedule.iterrows():
        program = row['Program']
        
        # Initialize a new calendar for each program if it doesn't exist
        if program not in programs:
            programs[program] = Calendar()
        
        event = Event()
        if row['Role'] == 'CURIOUS1' or row['Role'] == 'CURIOUS2':  
            event.name = "Curious Case"
        else:    
            event.name = f"{program} - Assessment"
        
        # Set event date
        event_date = datetime.strptime(row['Date'], '%Y-%m-%d').date()
        
        # Extract start and end times from the Time Slot column
        time_slot = row['Time Slot']
        start_time_str, end_time_str = time_slot.split(' - ')
        
        # Combine date with start and end times to create full datetime objects
        start_time = timezone.localize(datetime.combine(event_date, datetime.strptime(start_time_str.strip(), '%H:%M').time()))
        end_time = timezone.localize(datetime.combine(event_date, datetime.strptime(end_time_str.strip(), '%H:%M').time()))
        
        # Set the event start and end times
        event.begin = start_time
        event.end = end_time
        
        organizer_email = "pepijn.alofs@ormittalent.be"
        event.organizer = organizer_email
        
        # Add event description based on the Role
        if row['Role'] == 'CURIOUS1' or row['Role'] == 'CURIOUS2':
            event.description = "Curious Case"
        elif row['Role'] == 'CASE1' or row['Role'] == 'CASE2':
            event.description = f"Assessment {program} - Case Study"
        elif row['Role'] == 'DATACASE':
            event.description = f"Assessment {program} - Datacase"
        elif row['Role'] == 'ROLEPLAY1':
            event.description = f"Assessment {program} - Roleplay"        
        elif row['Role'] == 'PAPI1':
            event.description = f"Assessment {program} - PAPI"
        else:
            print('unique ' + row['Role'])
        
        assessor = row['Assessor']
        if assessor in emails:
            email_address = emails[assessor]
            attendee = Attendee(email=email_address, common_name=assessor, rsvp="TRUE")
            event.attendees.add(attendee)
        
        # Add the event to the corresponding program calendar
        programs[program].events.add(event)

    # Create output directory if it doesn't exist
    output_dir = os.path.join(base_path, 'Outlook Calendar Files')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Save each program's calendar to a separate ICS file
    for program, calendar in programs.items():
        ics_file_path = os.path.join(output_dir, f'{program.replace(" ", "_")}_Schedule_{startdate}-{enddate}.ics')
        with open(ics_file_path, 'w') as my_file:
            my_file.writelines(calendar)
        
        print(f"ICS file created for {program}: {ics_file_path}")


### Preparing the data ###
def makeSchedule(startDate, endDate, assessorExcel, output_text, check_calender=False, constant_goal_weight=1, want_ics=False):    
    if check_calender:
        log_message('Using calender availability for scheduling', output_text)
    else:
        log_message('NOT using calender availability for scheduling', output_text)
    
    log_message('Start scheduling...', output_text)

    # Define the time slots for each activity type
    time_slots = {
        'ROLEPLAY1': "12:00 - 16:00",  # General assessments
        'DATACASE': "12:00 - 16:00",   # General assessments
        'PAPI1': "12:00 - 16:00",      # General assessments
        'CASE1': "12:00 - 16:00",      # General assessments
        'CURIOUS1': {
            'Monday': "12:00 - 13:30",
            'Tuesday': "17:30 - 19:00",
            'Thursday': "17:30 - 19:00",
            'Friday': "09:00 - 10:30"
        },
        'CURIOUS2': {
            'Monday': "12:00 - 13:30",
            'Tuesday': "17:30 - 19:00",
            'Thursday': "17:30 - 19:00",
            'Friday': "09:00 - 10:30"
        }
    }
           
    datesResult = workingDays(startDate, endDate)
    dates = datesResult['workingDates']

    datesTuples = datesResult['workingDatesWithWeekdays']
    
    weeks = datesResult['workingWeeks']
    months = datesResult['workingMonths']

    assessors, program_capacities, officeUnavailabilities = load_data(assessorExcel)
    
    assessments = {
        "Curious": {"Activities": ["CURIOUS1", "CURIOUS2"], "Candidates": 1}, #No no of candidates
        "MCP&DATA": {"Activities": ["ROLEPLAY1", "DATACASE","PAPI1"], "Candidates": 3},
        "AM IT": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Buildwise": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Scrum Master": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Pluxee": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        ### Future Programs (Can be added by user by simply changing Progam's monthly goal and adding "ProgramX" label to assessors)
        "Program1": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Program2": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Program3": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Program4": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Program5": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Program6": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Program7": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
        "Program8": {"Activities": ["ROLEPLAY1", "CASE1","PAPI1"], "Candidates": 3},
    }
    
    # Get focus programs for each assessor based on programs in input sheet
    available_programs = {name: info['programs'] for name, info in assessors.items()}
    
    #Possible roles during assessment activities:
    assessmentActivities = {
        "CURIOUS1": {"Assessors": [name for name, info in assessors.items() if "CURIOUS" in info['Activities']], "Capacity": 3},
        "CURIOUS2": {"Assessors": [name for name, info in assessors.items() if "CURIOUS" in info['Activities']], "Capacity": 3},
        "ROLEPLAY1": {"Assessors": [name for name, info in assessors.items() if "ROLEPLAY" in info['Activities']], "Capacity": 5},
        "CASE1": {"Assessors": [name for name, info in assessors.items() if "CASE" in info['Activities']], "Capacity": 5},
        "PAPI1": {"Assessors": [name for name, info in assessors.items() if "PAPI" in info['Activities']], "Capacity": 9},
        "DATACASE": {"Assessors": [name for name, info in assessors.items() if "CASE" in info['Activities'] and info['DATA']], "Capacity": 5}
    }
    
    ### The model ###
    model = cp_model.CpModel()


    ## Create variables: All possible schedule options
    assignments = {}
    for date in dates:
        for assessmentType, assessmentInfo in assessments.items():
            for activity in assessmentInfo['Activities']:
                # Only consider assessors whose programs allows them to participate
                for assessor in assessmentActivities[activity]['Assessors']:
                    if assessor in available_programs and assessmentType in available_programs[assessor]:
                        varName = f'{date}_{assessmentType}_{activity}_{assessor}'
                        assignments[varName] = model.NewBoolVar(varName)
           

    # No assessment on 2 consecutive weekdays for assessors (except external)
    for date_idx, today in enumerate(datesResult['workingDates']):
        if date_idx + 1 < len(datesResult['workingDates']):
            tomorrow = datesResult['workingDates'][date_idx + 1]  # Get the next day
            
            for assessor in assessors.keys():
                if assessor != 'External':  # Skip External assessors
                    consecutiveday_count = 0  # Initialize the activity counter for this assessor
                    
                    # Loop through today's activities
                    for assessmentType_today, assessmentInfo in assessments.items():
                        for activity_today in assessmentInfo['Activities']:
                            if assessor in assessmentActivities[activity_today]['Assessors']:
                                today_label = f'{today}_{assessmentType_today}_{activity_today}_{assessor}'
                                
                                # If the assessor is assigned today, count the assignment
                                if today_label in assignments:
                                    consecutiveday_count += assignments[today_label]
                    
                    # Loop through tomorrow's activities
                    for assessmentType_tomorrow, assessmentInfo in assessments.items():
                        for activity_tomorrow in assessmentInfo['Activities']:
                            if assessor in assessmentActivities[activity_tomorrow]['Assessors']:
                                tomorrow_label = f'{tomorrow}_{assessmentType_tomorrow}_{activity_tomorrow}_{assessor}'
                                
                                # If the assessor is assigned tomorrow, count the assignment
                                if tomorrow_label in assignments:
                                    consecutiveday_count += assignments[tomorrow_label]
                    
                    # After checking both days' assignments, apply the constraint
                    model.Add(consecutiveday_count <= 1)  # Limit total activities to 1 across both days

    # Only 1 activity type per assessment, except for curious (i.e. no more than 1 PAPI, 1 (DATA)Case, 1 Roleplay and 2x CURIOUS per assignment type)
    for date in dates:
        for assessmentType, assessmentInfo in assessments.items():
            for activity in assessmentInfo['Activities']:
                if activity.startswith("CURIOUS"):  # Check if it's a Curious case
                    # No uniqueness constraint for Curious cases because you should have CURIOUS1 and CURIOUS2 at the same time
                    continue
                else:
                    model.Add(
                        sum(
                            assignments[f'{date}_{assessmentType}_{activity}_{assessor}']
                            for assessor in assessmentActivities[activity]['Assessors']
                            if f'{date}_{assessmentType}_{activity}_{assessor}' in assignments
                        ) <= 1
                    )

    # Every assessor only does 1 activity per day
    for date in dates:
            for assessor in assessors.keys():
                if assessor != 'External':
                    assessorCount = 0
                    for assessmentType, assessmentInfo in assessments.items():
                        for activity in assessmentInfo['Activities']:
                            if assessor in assessmentActivities[activity]['Assessors']:
                                key = f'{date}_{assessmentType}_{activity}_{assessor}'
                                if key in assignments:
                                    assessorCount += assignments[f'{date}_{assessmentType}_{activity}_{assessor}']
                    model.Add(assessorCount <= 1) 

    # Assessment afternoons: All activities of an assesment type (except curious case) should be scheduled together
    for assessmentType, assessmentInfo in assessments.items():
        assessmentLength = len(assessmentInfo['Activities'])
        for date in dates:
            activityVars = []
            for activity in assessmentInfo['Activities']:
                for assessor in assessmentActivities[activity]['Assessors']:
                    key = f'{date}_{assessmentType}_{activity}_{assessor}'
                    if key in assignments:
                        activityVars.append(assignments[f'{date}_{assessmentType}_{activity}_{assessor}'])

            # Constraint to ensure all activities are scheduled together or none are scheduled
            all_or_none = model.NewBoolVar(f'all_or_none_{date}_{assessmentType}')
            model.Add(sum(activityVars) == assessmentLength).OnlyEnforceIf(all_or_none)
            model.Add(sum(activityVars) == 0).OnlyEnforceIf(all_or_none.Not())


    # Capacity constraint: Assessors have a personal monthly capacity that shouldn't be exceeded
    for assessor, assessorInfo in assessors.items():
        for month, monthDates in months.items():
            model.Add(
            sum(
                assignments[f'{date}_{assessmentType}_{activity}_{assessor}'] * assessmentActivities[activity]['Capacity']
                for date in monthDates
                for assessmentType, assessmentInfo in assessments.items()
                for activity in assessmentInfo['Activities']
                if assessor in assessmentActivities[activity]['Assessors']
                and f'{date}_{assessmentType}_{activity}_{assessor}' in assignments
            ) <= assessorInfo["Capacity"][month]
        )

    # Don't schedule assessors if they have weekly non-working days (format: 0 through 4 for all weekdays, seperated by comma's)
    for assessor, assessorInfo in assessors.items():
        if "weeklyUnavailability" in assessorInfo:
            for unavailableDay in assessorInfo['weeklyUnavailability']:
                for date, weekday in datesTuples:
                    if weekday == unavailableDay:
                        for assessmentType, assessmentInfo in assessments.items():
                            for activity in assessmentInfo['Activities']:
                                if assessor in assessmentActivities[activity]['Assessors']:
                                    key = f'{date}_{assessmentType}_{activity}_{assessor}'
                                    if key in assignments:
                                        # Check if the assessor is 'Laetitia', the day is Friday, and the assessment is not a curious case
                                        if assessor != 'Laetitia':
                                            # Apply the unavailability constraint only for assessment afternoons
                                            model.Add(assignments[f'{date}_{assessmentType}_{activity}_{assessor}'] == 0)
                                        elif assessor == 'Laetitia':
                                            if weekday != 4:
                                                model.Add(assignments[f'{date}_{assessmentType}_{activity}_{assessor}'] == 0)
                                            else:
                                                if assessmentType != 'Curious':
                                                    model.Add(assignments[f'{date}_{assessmentType}_{activity}_{assessor}'] == 0)

    # If they have anything in the unavailability column (trainings for example), don't schedule them (overlap with new calender function!)
    for assessor, assessorInfo in assessors.items():
        if "Unavailability" in assessorInfo:
            for unavailableDate in assessorInfo['Unavailability']:
                for assessmentType, assessmentInfo in assessments.items():
                    for activity in assessmentInfo['Activities']:
                        if assessor in assessmentActivities[activity]['Assessors']:
                            model.Add(assignments[varName] == 0)

    # If the office is unavailable, don't plan anything 
    for unavailableDate in officeUnavailabilities:
        for assessmentType, assessmentInfo in assessments.items():
            for activity in assessmentInfo['Activities']:
                for assessor in assessmentActivities[activity]['Assessors']:
                    varName = f'{unavailableDate}_{assessmentType}_{activity}_{assessor}'
                    if varName in assignments:  # Check if the variable exists
                        model.Add(assignments[varName] == 0)
   
    if check_calender:
        for assessor, assessor_info in assessors.items():
            if assessor == 'External':
                continue  # Skip to next assessor, since 'External' is always available
            
            assessment_dates = set(assessor_info.get('assessmentAvailability', []))  # Use set for faster lookups
            case_dates = set(assessor_info.get('caseAvailability', []))  # Use set for faster lookups
            
            for date in dates:  # Go over all dates
                # Apply hard constraints based on availability
                for assessment_type, assessment_info in assessments.items():
                    for activity in assessment_info['Activities']:
                        key = f'{date}_{assessment_type}_{activity}_{assessor}'
                        if key in assignments:
                            # If it's a regular assessment activity (not Curious), check for assessment availability
                            if activity not in ["CURIOUS1", "CURIOUS2"]:
                                # Only schedule if date is in assessment availability
                                if date not in assessment_dates:
                                    model.Add(assignments[key] == 0)  # Not available for assessment
                            else:
                                # Handle Curious cases (assumed to be case-specific availability)
                                if date not in case_dates:
                                    model.Add(assignments[key] == 0)  # Not available for Curious case

            # # Print results for the assessor
            # print(f"Assessor: {assessor}")
            # print(f"Assessment Available Dates: {assessment_dates}")
            # print(f"Case Available Dates: {case_dates}")
            
    # Make sure curious case are on monday, friday and mutually exclusive tuesday/thursday
    tuesday_cases = {}
    thursday_cases = {}
    # Loop through the dates and weekdays
    for date, weekday in datesTuples:
        if date in officeUnavailabilities:
            continue  # Skip office unavailability days
    
        for activity in ["CURIOUS1", "CURIOUS2"]:
            # Allow Curious cases but don't require them on Monday, Tuesday, Thursday, and Friday
            if weekday == 0 or weekday == 1 or weekday == 3 or weekday == 4:  # Monday, Tuesday, Thursday, Friday
                model.Add(sum(assignments[f'{date}_Curious_{activity}_{assessor}'] for assessor in assessmentActivities[activity]['Assessors']) <= 1)
                
                # Track Tuesday and Thursday Curious case assignments
                if weekday == 1:  # Tuesday
                    tuesday_cases[date] = sum(assignments[f'{date}_Curious_{activity}_{assessor}'] for assessor in assessmentActivities[activity]['Assessors'])
                elif weekday == 3:  # Thursday
                    thursday_cases[date] = sum(assignments[f'{date}_Curious_{activity}_{assessor}'] for assessor in assessmentActivities[activity]['Assessors'])
    
            # Disallow Curious cases on Wednesday
            elif weekday == 2:  # Wednesday
                model.Add(sum(assignments[f'{date}_Curious_{activity}_{assessor}'] for assessor in assessmentActivities[activity]['Assessors']) == 0)
    
    # Mutual exclusion for Tuesday and Thursday
    for tuesday_date, tuesday_assignment in tuesday_cases.items():
        # Find the Thursday of the same week
        tuesday_date = pd.to_datetime(tuesday_date, format='%Y-%m-%d')  # or the appropriate format of your dates
        thursday_date = tuesday_date + pd.DateOffset(days=2)        
        # If the corresponding Thursday exists, add the mutual exclusion constraint
        if thursday_date in thursday_cases:
            model.Add(tuesday_assignment + thursday_cases[thursday_date] <= 1)

    # Maximum 2 activities per week for each assessor, 2 cases | 1 case & 1 assessment day | NOT 2 assessment days (too intense)
    for weekNum, weekDates in weeks.items():
        for assessor in assessors.keys():
            if assessor != 'External':
                weeklyActivityCount = 0
                assessmentDayCount = 0  # Counter for assessment days
                for date in weekDates:
                    for assessmentType, assessmentInfo in assessments.items():
                        for activity in assessmentInfo['Activities']:
                            if assessor in assessmentActivities[activity]['Assessors']:
                                key = f'{date}_{assessmentType}_{activity}_{assessor}'
                                if key in assignments:
                                    weeklyActivityCount += assignments[f'{date}_{assessmentType}_{activity}_{assessor}']
                                    # Check if the activity is an assessment day activity
                                    if activity not in ['CURIOUS1', 'CURIOUS2']:
                                        assessmentDayCount += assignments[key]
                model.Add(weeklyActivityCount <= 2)
                model.Add(assessmentDayCount <= 1)  # Restrict to only 1 assessment day per week
            
    #Ensure at least one HR team member per assessment
    ## with the current capacity it's not possible to find a solution with this constraint active...
    if False:  
        for date in weekDates:
            for  assessmentType, assessmentInfo in assessments.items():
                key = f'{date}_{assessmentType}_{activity}_{assessor}'
                if key in assignments:
                    model.Add( sum(assignments[f'{date}_{assessmentType}_{activity}_{assessor}']
                                for activity in assessmentInfo['Activities']
                                for assessor in assessmentActivities[activity]['Assessors'] if assessors[assessor]['HR']) >=1 )
   
    external_assessor_count = sum(
    assignments[f'{date}_{assessmentType}_{activity}_External']
    for date in dates
    for assessmentType, assessmentInfo in assessments.items()
    for activity in assessmentInfo['Activities']
    if 'External' in assessmentActivities[activity]['Assessors']
    and f'{date}_{assessmentType}_{activity}_External' in assignments  # Ensure key exists
)
    
    # Count how many sessions have not been planned (underscheduled)
    goal_deviations = []
    for month, monthDates in months.items():
        month_name = get_month_name(month)  # Convert month number to name
        for program in assessments.keys():
            program_goal = program_capacities[month_name][program]  # Get the program goal for the month
        
            total_candidates_for_program = 0
            for date in monthDates:
                for assessmentType, assessmentInfo in assessments.items():
                    if program in assessmentType:
                        activity = assessmentInfo['Activities'][0]  # Assuming first activity for scheduling
                        for assessor in assessmentActivities[activity]['Assessors']:
                            key = f'{date}_{assessmentType}_{activity}_{assessor}'
                            if key in assignments:
                                total_candidates_for_program += assignments[key] * assessments[assessmentType]['Candidates']
            
            # Ensure the total number of scheduled candidates doesn't exceed the program goal
            under_goal = model.NewIntVar(0, program_goal, f'under_goal_{month}_{program}')
            
            # Allow under-scheduling but prevent over-scheduling
            model.Add(total_candidates_for_program <= program_goal)
            model.Add(program_goal - total_candidates_for_program <= under_goal)  # Only under-achievement allowed
            
            # Add the under_goal to the deviations for minimization
            goal_deviations.append(under_goal)

    # Update the objective function to minimize both external assessor usage and goal deviations
    model.Minimize(external_assessor_count + constant_goal_weight * sum(goal_deviations))

    ### The Solution ###
    solver = cp_model.CpSolver()
    
    #NEW: 
    days_difference = (endDate - startDate).days
    print(f"Difference between start and end: {days_difference} days - Time limit: {120 + 2*days_difference} sec.")
    solver.parameters.max_time_in_seconds = 120 + 2*days_difference  # Set the maximum time limit depending on date range
    
    solver.parameters.log_search_progress = False
    status = solver.Solve(model)

    solutionDf = pd.DataFrame(columns=['Date', 'Program', 'Role', 'Total Capacity Cost', 'Assessor', 'Time Slot'])
    
    
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        log_message('Scheduling finished, compiling file...', output_text)
        scheduleText = ""
        for date in dates:
            scheduleText += f"Date: {date}\n"
            weekday = pd.to_datetime(date, format='%Y-%m-%d').strftime('%A')  # Get the weekday name
    
            for assessmentType, assessmentInfo in assessments.items():
                for activity in assessmentInfo['Activities']:
                    for assessor in assessmentActivities[activity]['Assessors']:
                        if f'{date}_{assessmentType}_{activity}_{assessor}' in assignments:
                            if solver.Value(assignments[f'{date}_{assessmentType}_{activity}_{assessor}']):
                                # Determine the time slot based on the activity and weekday
                                if isinstance(time_slots[activity], dict):  # For cases and Curious, times depend on the day
                                    time_slot = time_slots[activity].get(weekday, "Unknown Time")
                                else:
                                    time_slot = time_slots[activity]  # For general assessments
    
                                if assessmentType == 'Curious':
                                    scheduleText += f"Curious case scheduled with {assessor} at {time_slot}\n"
                                    program = 'Curious Case'
                                else:
                                    scheduleText += f"Assessment afternoon scheduled for {assessmentInfo['Candidates']} candidates with {assessor} assigned to {activity} at {time_slot}\n"
                                    for prog in ['MCP&DATA', 'AM IT', 'Buildwise', 'Scrum Master', 'Pluxee', 'Program1', 'Program2', 'Program3', 'Program4', 'Program5', 'Program6', 'Program7', 'Program8']:
                                        if prog in assessmentType:
                                            program = prog
                                            break
                                            
                                # Append to the DataFrame, including the time slot
                                capacityCost = assessmentActivities[activity]['Capacity']
                                new_row = pd.DataFrame([{
                                    'Date': date, 
                                    'Month': pd.to_datetime(date, format='%Y-%m-%d').month,
                                    'Program': program,
                                    'Role': activity, 
                                    'Total Capacity Cost': capacityCost, 
                                    'Assessor': assessor,
                                    'Time Slot': time_slot
                                }])
                                solutionDf = pd.concat([solutionDf, new_row], ignore_index=True)


        solutionDf['Date'] = pd.to_datetime(solutionDf['Date'], format='%Y-%m-%d')
        solutionDf['Month'] = solutionDf['Date'].dt.month
        capacityUsage = solutionDf.groupby(['Month', 'Assessor'])['Total Capacity Cost'].sum().reset_index()
        capacityUsage['Total Capacity'] = capacityUsage.apply(lambda row: assessors[row['Assessor']]['Capacity'][row['Month']], axis=1)
        # Insert two empty columns before 'Total Capacity'
        capacityUsage.insert(capacityUsage.columns.get_loc('Total Capacity'), 'Curious Case', '')
        capacityUsage.insert(capacityUsage.columns.get_loc('Total Capacity'), 'PAPI', '')
        capacityUsage.insert(capacityUsage.columns.get_loc('Total Capacity'), 'Roleplay', '')
        capacityUsage.insert(capacityUsage.columns.get_loc('Total Capacity'), 'Business Case', '')
        capacityUsage.insert(capacityUsage.columns.get_loc('Total Capacity'), 'Datacase', '')
        capacityUsage['Remaining Capacity'] = capacityUsage['Total Capacity'] - capacityUsage['Total Capacity Cost']
        #V2: Also add the ones who weren't scheduled
        for month in solutionDf['Month'].unique():
            filtered_df = solutionDf[solutionDf['Month'] == month]
            assessors_this_month = filtered_df['Assessor'].unique()
            for assessor in assessors.keys():
                if assessor not in assessors_this_month:
                    print(assessors[assessor]['Capacity'][month])
                    new_row_df = pd.DataFrame([{'Month': month, 'Assessor': assessor, 'Curious Case': 0, 'PAPI': 0, 'Roleplay': 0, 'Business Case': 0, 'Datacase': 0, 'Total Capacity Cost': 0, 'Total Capacity': assessors[assessor]['Capacity'][month], 'Remaining Capacity': assessors[assessor]['Capacity'][month]}])
                    capacityUsage = pd.concat([capacityUsage, new_row_df], ignore_index=True)
                    
        # Sort the DataFrame first by 'Month', then by 'Total Capacity Cost'
        # Sort by 'Month' ascending and then by 'Total Capacity Cost' descending
        capacityUsage = capacityUsage.sort_values(by=['Month', 'Total Capacity Cost'], ascending=[True, False])

        
        solutionDf['Date'] = solutionDf['Date'].dt.strftime('%Y-%m-%d') #Make sure not to include time
        solutionDf = solutionDf[['Date', 'Time Slot', 'Program', 'Role', 'Total Capacity Cost', 'Assessor', 'Month']] #Rearrange order

        ##NEW
        #Initialize a new DataFrame for goal comparison
        goal_comparison_df = pd.DataFrame(columns=['Month', 'Program', 'Initial Goal', 'Final Scheduled'])
        
        # Iterate through each month and program to compare initial goals with the final result
        for month, monthDates in months.items():
            month_name = get_month_name(month)  # Convert month number to name
            for program in assessments.keys():
                program_goal = program_capacities[month_name][program]  # Get the program goal for the month
            
                # Initialize a variable to count scheduled candidates for this program
                total_candidates_for_program = 0
                for date in monthDates:
                    if program == 'Curious':  # Special case handling for Curious program
                        curious_assigned = False  # Track if a Curious case is counted already
                        for assessor in assessmentActivities['CURIOUS1']['Assessors']:
                            key1 = f'{date}_Curious_CURIOUS1_{assessor}'
                            key2 = f'{date}_Curious_CURIOUS2_{assessor}'
                            
                            if key1 in assignments and key2 in assignments and not curious_assigned:
                                # If both CURIOUS1 and CURIOUS2 are assigned, count it as one case
                                if solver.Value(assignments[key1]) or solver.Value(assignments[key2]):
                                    total_candidates_for_program += 1  # Count the curious case as one
                                    curious_assigned = True
                    else:
                        for assessmentType, assessmentInfo in assessments.items():
                            if program in assessmentType:  # Match program to the correct assessment type
                                activity = assessmentInfo['Activities'][0]  # Assuming you want the first activity for scheduling
                                for assessor in assessmentActivities[activity]['Assessors']:
                                    key = f'{date}_{assessmentType}_{activity}_{assessor}'
                                    if key in assignments:
                                        total_candidates_for_program += solver.Value(assignments[key]) * assessments[assessmentType]['Candidates']
                
                # Append the results to the DataFrame
                candidates_per_program = assessments.get(program, {}).get('Candidates', 1)  # Default to 1 if not found
                new_row = pd.DataFrame([{
                    'Month': month,
                    'Program': program,
                    'Initial Goal': program_capacities[month_name][program]/candidates_per_program,
                    'Final Scheduled': total_candidates_for_program/candidates_per_program,
                    'Difference': (total_candidates_for_program/candidates_per_program)-(program_capacities[month_name][program]/candidates_per_program)
                }]) 
                goal_comparison_df = pd.concat([goal_comparison_df, new_row], ignore_index=True)
              
        # Adjust the credit per activity based on the activity type (Curious = 3, General Assessments = 5, PAPI1 = 9)
        # Update how initial and final credits are calculated based on the program type
        goal_comparison_df['Initial Credits'] = goal_comparison_df.apply(
            lambda row: row['Initial Goal'] * (6 if row['Program'] == 'Curious' else 19), axis=1 #5 + 5 + 9
        )
        
        goal_comparison_df['Final Credits'] = goal_comparison_df.apply(
            lambda row: row['Final Scheduled'] * (6 if row['Program'] == 'Curious' else 19), axis=1
        )

        # Split the data between Curious cases and other sessions
        curious_df = goal_comparison_df[goal_comparison_df['Program'] == 'Curious']
        non_curious_df = goal_comparison_df[goal_comparison_df['Program'] != 'Curious']
        
        # Calculate total initial and final credits for Curious and non-Curious
        total_initial_credits_curious = curious_df['Initial Credits'].sum()
        total_final_credits_curious = curious_df['Final Credits'].sum()
        
        total_initial_credits_non_curious = non_curious_df['Initial Credits'].sum()
        total_final_credits_non_curious = non_curious_df['Final Credits'].sum()
        
        # Calculate the success percentages for Curious and non-Curious cases
        credits_success_percentage_curious = (total_final_credits_curious / total_initial_credits_curious) * 100 if total_initial_credits_curious > 0 else 0
        credits_success_percentage_non_curious = (total_final_credits_non_curious / total_initial_credits_non_curious) * 100 if total_initial_credits_non_curious > 0 else 0
        credits_success_percentage = ((total_final_credits_curious + total_final_credits_non_curious)  / (total_initial_credits_curious + total_initial_credits_non_curious)) * 100 if (total_initial_credits_curious + total_initial_credits_non_curious) > 0 else 0

        # If indicated to create ICS files, check success and call makeICS
        if credits_success_percentage >= 90 and want_ics:          
            # Check if solutionDf is empty before calling makeICS
            if not solutionDf.empty:
                makeICS(solutionDf, startDate, endDate)
            else:
                print("The solution DataFrame is empty. No ICS files will be created.")
                    
        # Create a summary row with the totals and both percentages
        summary_row = pd.DataFrame([{
            'Month': '',  # Empty for the summary row
            'Program': '',  # Empty for the summary row
            'Initial Goal': '',
            'Final Scheduled': '', 
            'Difference': '',
            '% Assessment Planned': round(credits_success_percentage, 2),
            '% Curious': round(credits_success_percentage_curious, 2),
            '% Assessment Afternoons': round(credits_success_percentage_non_curious, 2)            
        }])

        # Delete any programs where for that month the goal was 0 
        goal_comparison_df = goal_comparison_df[goal_comparison_df['Initial Goal'] > 0]
                
        # Append the summary row to the goal_comparison_df
        goal_comparison_df = pd.concat([goal_comparison_df, summary_row], ignore_index=True)
        
        # Optional: Adjust columns order if you want both percentages and assessor-sessions to be part of the output
        goal_comparison_df = goal_comparison_df[['Month', 'Program', 'Initial Goal', 'Final Scheduled', 'Difference', '% Assessment Planned', '% Curious', '% Assessment Afternoons']]

    else: #No solution
        solutionDf = pd.DataFrame()
        capacityUsage = pd.DataFrame()
        goal_comparison_df = pd.DataFrame()
        scheduleText = "No solution found. \n"

    return solutionDf, capacityUsage, goal_comparison_df, scheduleText