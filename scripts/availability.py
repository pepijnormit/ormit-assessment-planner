import json
import pandas as pd
import os
import sys
import requests
from datetime import datetime, time, timedelta
from ics import Calendar
import re
from tkinter import filedialog, END
import icalendar
import recurring_ical_events
import pytz

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
    
def clean_timezone_ics(ics_text):
    """
    Clean up invalid TZID entries in some .ics files
    """
    return re.sub(r'TZID:[^\r\n]*', 'TZID:Romance Standard Time', ics_text)

def merge_overlapping_events(df):
    # Ensure the data is sorted by start date and start time for proper merging
    df = df.sort_values(by=["Start Date", "Start Time"]).reset_index(drop=True)

    merged_events = []
    current_event = df.iloc[0]

    for i in range(1, len(df)):
        next_event = df.iloc[i]

        # Check if the events overlap (same start date and overlapping time)
        if (current_event["Start Date"] == next_event["Start Date"] and
            current_event["End Time"] >= next_event["Start Time"]):
            # Merge the events by extending the end time to the max end time
            current_event["End Time"] = max(current_event["End Time"], next_event["End Time"])
        else:
            # No overlap, append the current event and start a new one
            merged_events.append(current_event)
            current_event = next_event

    # Append the last event
    merged_events.append(current_event)
    merged_df = pd.DataFrame(merged_events)
    return merged_df

def get_calender(ics_url, staff_name, start_date_string, end_date_string):
    response = requests.get(ics_url[staff_name])
    
    if response.status_code == 200:
        print(f'Calendar {staff_name} successfully accessed')
        
        start_date = datetime.strptime(start_date_string, "%Y-%m-%d")
        paris_tz = pytz.timezone('Europe/Paris')
        start_date = paris_tz.localize(start_date)
        
        end_date = datetime.strptime(end_date_string, "%Y-%m-%d")
        paris_tz = pytz.timezone('Europe/Paris')
        end_date = paris_tz.localize(end_date)
           
        ical_string = response.text  # Get the content of the ICS file as a string
        a_calendar = icalendar.Calendar.from_ical(ical_string)
    
        # Use recurring_ical_events to find events within the specified range
        events = recurring_ical_events.of(a_calendar).between(start_date, end_date)
    
        # Store event details
        event_names = []
        start_dates = []
        end_dates = []
        # V2
        statuses = []
    
        # Collect event details
        for event in events:
            start = event["DTSTART"].dt  # Get the start time of the event
            end = event["DTEND"].dt  # Get the end time of the event
            summary = event["SUMMARY"]  # Get the summary (title) of the event
            busystatus = event["X-MICROSOFT-CDO-BUSYSTATUS"] #V2
    
            event_names.append(summary)
            start_dates.append(start)
            end_dates.append(end)
            statuses.append(busystatus)
    
        # Create the DataFrame with event details
        event_data = {
            "Event Name": event_names,
            "Start Date": [start.date() if isinstance(start, datetime) else start for start in start_dates],  # Handle datetime and ensure just date
            "Start Time": [start.strftime("%H:%M:%S") for start in start_dates],  # Format time as HH:MM:SS
            "End Date": [end.date() if isinstance(end, datetime) else end for end in end_dates],  # Handle datetime and ensure just date
            "End Time": [end.strftime("%H:%M:%S") for end in end_dates],  # Format time as HH:MM:SS
            "Status": statuses #V2
        }
    
        # Create a pandas DataFrame
        event_df = pd.DataFrame(event_data)
    
        # Ensure that Start Date and Start Time are properly sorted
        event_df['Start Date'] = pd.to_datetime(event_df['Start Date'])
        event_df['Start Time'] = pd.to_datetime(event_df['Start Time'], format="%H:%M:%S").dt.time
    
        # Sort the DataFrame by Start Date and then by Start Time
        event_df = event_df.sort_values(by=['Start Date', 'Start Time'])
    
        # Create a list to store expanded rows
        expanded_rows = []
        end_of_day_time = "23:59:59"
        start_of_day_time = "00:00:00"
        
        # Iterate over each event and generate rows for each day between the start and end dates
        for _, row in event_df.iterrows():
            event_dates = pd.date_range(row['Start Date'], row['End Date'], freq='D').date
    
            for i, event_date in enumerate(event_dates):
                # Determine the appropriate start and end times for each day
                if i == 0:  # First day of the event
                    start_time = str(row['Start Time'])
                    end_time = end_of_day_time if len(event_dates) > 1 else row['End Time']
                elif i == len(event_dates) - 1:  # Last day of the event
                    start_time = start_of_day_time
                    end_time = row['End Time']
                else:  # Middle days for multi-day events
                    start_time = start_of_day_time
                    end_time = end_of_day_time
                # end_time = row['End Time'] if i == len(event_dates) - 1 else end_of_day_time
                expanded_rows.append({
                    "Event Name": row['Event Name'],
                    "Start Date": event_date,
                    "Start Time": datetime.strptime(start_time, "%H:%M:%S").time(),
                    "End Date": event_date,
                    "End Time": datetime.strptime(end_time, "%H:%M:%S").time(),
                    "Status": row['Status'] #V2
                })
    
        # Create the expanded DataFrame
        expanded_event_df = pd.DataFrame(expanded_rows)
    
        # Sort and reset index
        expanded_event_df['Start Date'] = pd.to_datetime(expanded_event_df['Start Date']).dt.date
        expanded_event_df['Start Time'] = pd.to_datetime(expanded_event_df['Start Time'], format="%H:%M:%S").dt.time
        expanded_event_df = expanded_event_df.sort_values(by=['Start Date', 'Start Time']).reset_index(drop=True)
    
        #Before taking out 'next day' left-overs
        print(expanded_event_df)
        
        condition = (expanded_event_df['Start Date'] == expanded_event_df['End Date']) & (expanded_event_df['Start Time'] == expanded_event_df['End Time'])
        df = expanded_event_df[~condition]
        
        # Drop rows where the Status is 'FREE' V2
        df = df[df['Status'] != 'FREE']
        df['Event Name'] = df['Event Name'].str.slice(0, 7)  
        print(df)

        #Blur overlapping events into 1 big event
        df_merged = merge_overlapping_events(df)
        print(df_merged)
        return df_merged
    else:
        print(f"Error fetching calendar for {staff_name}: {response.status_code}")
        return pd.DataFrame()         

def get_free_time(schedule_df, staff_member, start_date_string, end_date_string):
    # Ensure that 'Start Date' and 'End Date' columns are in datetime format
    schedule_df['Start Date'] = pd.to_datetime(schedule_df['Start Date'])
    schedule_df['End Date'] = pd.to_datetime(schedule_df['End Date'])

    # Sort by the start time of events
    schedule_df = schedule_df.sort_values(by=['Start Date', 'Start Time'])

    start_date = datetime.strptime(start_date_string, "%Y-%m-%d")        
    end_date = datetime.strptime(end_date_string, "%Y-%m-%d")
    
    # Create a list to hold the result
    free_time_list = []

    # Process each day in the schedule
    day = start_date
    while day <= end_date:
        print(day.date())  # Print or process the current date
        events_for_day = schedule_df[schedule_df['Start Date'].dt.date == day.date()]
        print(events_for_day)
        if events_for_day.empty:
            print(events_for_day)
            free_time_list.append(['Free', day.date(), time(0, 0, 0), time(23, 59, 59), staff_member])
        else:
            # Track free time periods for the current day
            free_periods = []
            current_day_start = datetime.combine(day, datetime.min.time())
            day_end = datetime.combine(day, datetime.max.time())
    
            # Convert 'Start Time' and 'End Time' to datetime objects for accurate comparison
            first_event_start = datetime.combine(day, events_for_day.iloc[0]['Start Time'])
            last_event_end = datetime.combine(day, events_for_day.iloc[-1]['End Time'])
    
            # 1. Free time before the first event (if applicable)
            if first_event_start > current_day_start:
                free_periods.append((current_day_start, first_event_start))
    
            # 2. Free time between consecutive events (splitting periods when necessary)
            for i in range(1, len(events_for_day)):
                previous_event_end = datetime.combine(day, events_for_day.iloc[i - 1]['End Time'])
                current_event_start = datetime.combine(day, events_for_day.iloc[i]['Start Time'])
    
                # If there's a gap between events, register it as free time
                if current_event_start > previous_event_end:
                    free_periods.append((previous_event_end, current_event_start))
    
            # 3. Free time after the last event till the end of the day (if applicable)
            if last_event_end < day_end:
                free_periods.append((last_event_end, day_end))
    
            # 4. Add the identified free periods to the final list
            for period in free_periods:
                free_time_list.append(['Free', period[0].date(), period[0].time(), period[1].time(), staff_member])
        day += timedelta(days=1)  # Move to the next day
        
    # Convert the list into a DataFrame and return
    free_time_df = pd.DataFrame(free_time_list, columns=['Event Name', 'Start Date', 'Start Time', 'End Time', 'Staff'])
    return free_time_df

# def make_calenders(calenders, key, start_date, end_date, working_days):   
#     schedule = get_calender(calenders, key, str(start_date), str(end_date))
#     free_time = get_free_time(schedule, key)
#     return free_time

def find_assessment_slot(free_time_df, starttime, endtime, dayofweek):
    results = []

    for index, row in free_time_df.iterrows():
        start_time_free_slot = row['Start Time']
        end_time_free_slot = row['End Time']
        start_date = row['Start Date']
        
        if start_date.weekday() in dayofweek:
            start_time_event = pd.to_datetime(starttime).time()
            end_time_event = pd.to_datetime(endtime).time()

            # Check if the free slot accomodates (is at least starting and ending from) the start and end of the to-be-planned event
            if start_time_free_slot <= start_time_event and end_time_free_slot >= end_time_event:
                results.append(row['Start Date'])

    # Getting unique dates from results
    result = list(set(results))
    final_results = sorted(list(set(results)))
    return final_results

##### Data
def retrieve_calenders(template_file, output_text, start_date, end_date, links='resources/calenders.json'):
    # working_days = {
    #     "Monday": ["morning", "afternoon"],
    #     "Tuesday": ["morning", "afternoon"],
    #     "Wednesday": ["morning", "afternoon"],
    #     "Thursday": ["morning", "afternoon"],
    #     "Friday": ["morning", "afternoon"]
    # }
    
    calender_file = os.path.join(base_path, 'resources', 'calenders.json')
    with open(calender_file, 'r') as json_file:
        calenders = json.load(json_file)
            
    ### Add availability to the selected excel
    file_path = template_file
    excel_file = pd.read_excel(file_path, sheet_name=None)  # sheet_name=None loads all sheets
    #Loop over all calenders
    for key in excel_file.keys():   #For all assessors in assessor file
        if key != 'Extra' and key != 'External':
            if key in calenders.keys(): #If calender link is found
                schedule = get_calender(calenders, key, str(start_date), str(end_date))
                avail_gen = get_free_time(schedule, key, str(start_date), str(end_date)) #V2
                if avail_gen.empty == False: #pd.DataFrame([]) if a person has 0 free moments
                    avail_afternoon = find_assessment_slot(avail_gen, starttime="12:00:00", endtime="16:00:00", dayofweek=[0,1,2,3]) #Not Friday
                    avail_afternoon_str = ', '.join([dt.strftime('%Y-%m-%d') for dt in avail_afternoon])                  
                    avail_case1 = find_assessment_slot(avail_gen, starttime="12:00:00", endtime="13:30:00", dayofweek=[0]) #Monday
                    avail_case2 = find_assessment_slot(avail_gen, starttime="17:30:00", endtime="19:00:00", dayofweek=[1,3]) #Tuesday or Thursday
                    avail_case3 = find_assessment_slot(avail_gen, starttime="9:00:00", endtime="10:30:00", dayofweek=[4]) #Friday
                    combined_avail = avail_case1 + avail_case2 + avail_case3
                    avail_case = sorted(combined_avail)
                    avail_case_str = ', '.join([dt.strftime('%Y-%m-%d') for dt in avail_case])
                    log_message(f"{key} has {len(avail_afternoon)} slots available for assessment and {len(avail_case)} for cases between {start_date} and {end_date}", output_text)
                    
                    # Select the sheet you want to work with
                    df = excel_file[key] 
                    
                    # Find the index of the row where the first cell contains "Unavailability"
                    unavailability_row = df[df.iloc[:, 0] == 'Unavailability'].index[0]
                    
                    # Create new rows with specific content
                    new_row_1 = pd.DataFrame([['assessmentAvailability', avail_afternoon_str] + [''] * (len(df.columns) - 2)], columns=df.columns)
                    new_row_2 = pd.DataFrame([['caseAvailability', avail_case_str] + [''] * (len(df.columns) - 2)], columns=df.columns)
                    
                    # Insert the new rows above the row containing "Unavailability"
                    df_updated = pd.concat([df.iloc[:unavailability_row], new_row_1, new_row_2, df.iloc[unavailability_row:]]).reset_index(drop=True)
                    
                    # Replace the updated DataFrame in the dictionary
                    excel_file[key] = df_updated
                else:
                    log_message(f"{key} has 0 slots available for assessment and 0 for cases between {start_date} and {end_date}", output_text)

            else: #No calender link found for this assessor
                log_message(f"No calender found for {key}: Assume full availability", output_text) 
                all_dates = [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]
                date_list = [date.strftime("%Y-%m-%d") for date in all_dates]
                full_availability = ', '.join(date_list)
                
                df = excel_file[key] 
                new_row_1 = pd.DataFrame([['assessmentAvailability', full_availability] + [''] * (len(df.columns) - 2)], columns=df.columns)
                new_row_2 = pd.DataFrame([['caseAvailability', full_availability] + [''] * (len(df.columns) - 2)], columns=df.columns)
             
                # Find the index of the row where the first cell contains "Unavailability"
                unavailability_row = df[df.iloc[:, 0] == 'Unavailability'].index[0]
                # Insert the new rows above the row containing "Unavailability"
                df_updated = pd.concat([df.iloc[:unavailability_row], new_row_1, new_row_2, df.iloc[unavailability_row:]]).reset_index(drop=True)
                
                # Replace the updated DataFrame in the dictionary
                excel_file[key] = df_updated
        
    # Save the updated dictionary of DataFrames back to Excel
    output_file = os.path.join(base_path, 'resources', 'assessors2025_available.xlsx')
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for sheet_name, df_sheet in excel_file.items():
            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
    log_message(f"Availability succesfully exported to excel: {output_file}", output_text)  
    return output_file



#retrieve_calenders()