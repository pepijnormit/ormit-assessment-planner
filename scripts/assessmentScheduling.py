"""
V2 (labeled in comments in code):
* In functionScript, added not-scheduled assessors to Capacity Usage sheet in final schedule excel so it's easy to check reason for absence in schedule



V1
All dates are assumed to be: YYYY-MM-DD
----------
Input file:
*Staff member*: Sheet has name of staff member, example:
----------
Key	                    Value
Activities	            ROLEPLAY, CASE
HR	                    FALSE
DATA	                FALSE
programs	            MCP&DATA
weeklyUnavailability	
assessmentAvailability	2025-01-01, 2025-01-02, 2025-01-06 $Y-%m-%d !!!!
caseAvailability	    2025-01-02, 2025-01-03, 2025-01-06
Unavailability	
Capacity - January	    7.2
Capacity - February	    7.2
Capacity - March	    7.2
Capacity - April	    7.2
Capacity - May	        7.2
Capacity - June	        7.2
Capacity - July	        7.2
Capacity - August	    7.2
Capacity - September	7.2
Capacity - October	    7.2
Capacity - November	    7.2
Capacity - December	    7.2

*Extra*  sheet for candidate goals details
Note: For all progams 1 assessment day is x number of candidates (here: 3, so goal of 9 is 3 days). For curious cases, the goal is x1 (so 14 curious cases = 14 curious cases)
----------
Key	Value	              MCP&DATA	AM IT	Buildwise	Scrum Master	Pluxee	Curious
Public Holidays	          2025-01-01, 2025-04-21 
Office Events	          2024-04-15, 2024-04-16
Candidate Goal - January  9	    3	3	3	0	14
Candidate Goal - February 12	3	3	3	0	10
Candidate Goal - March	  15	0	3	0	3	10
Candidate Goal - April	  18	0	3	3	3	10
Candidate Goal - May	  15	3	3	3	6	11
Candidate Goal - June	  12	3	3	3	3	10
Candidate Goal - July	   9	0	3	0	3	11
Candidate Goal - August	  15	3	6	0	3	14
Candidate Goal - September 18	0	0	0	0	14
Candidate Goal - October   18	0	0	0	0	13
Candidate Goal - November   9	0	0	0	0	20
Candidate Goal - December	3	0	0	0	0	6
"""
import os
import sys
import customtkinter as ctk
import tkinter.font as tkFont
from tkinter import filedialog, END
import tkinter.messagebox as messagebox
from tkcalendar import Calendar
from functionScript import makeSchedule
from availability import retrieve_calenders
import pandas as pd
from datetime import datetime, date
import xlsxwriter #For active formulas in output sheet
import threading #To avoid non-responsive GUI

if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

def slider_callback(value):
    integer_value = round(value)  # Ensure it's an integer
    value_label.configure(text=f"Value: {value:.2f}")  # Update the label dynamically
    set_slider_color(integer_value)

def set_slider_color(value):
    # Define colors for three categories
    colors = [
        "#00FF00", "#00FF00", "#00FF00", "#00FF00", "#00FF00",  # Green 
        "#00FF00", "#00FF00", "#00FF00", "#00FF00", "#00FF00",  # Orange
        "#FFFF00", "#FFFF00", "#FFA500", "#FFA500", "#FF4500",  # Red 
    ]
    step_size = 3 / 15  # Each step covers 3/15 of the range
    index = min(int(value / step_size), 14)  # Calculate index (0–14)
    slider.configure(progress_color=colors[index])

def start_scheduling_threaded():
    # Start the scheduling process in a separate thread to avoid freezing of GUI
    scheduling_thread = threading.Thread(target=start_scheduling, daemon=True) #daemon prevents lingering threads background after closing
    scheduling_thread.start()
    
def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        file_entry_var.set(file_path)
        
def log_message(message):
    output_text.configure(state='normal')  # Enable editing
    output_text.insert(END, f"{message}\n")  # Insert message at the end
    output_text.configure(state='disabled')  # Disable editing after update
    output_text.yview(END)  # Scroll to the latest message
    output_text.update_idletasks()  # Force the GUI to update immediately

def start_scheduling():
    selectedFile = file_entry.get()

    log_message(f"Selected file path: '{selectedFile}'")
    if os.path.exists(selectedFile):
        try:
            excel_file = pd.read_excel(selectedFile, sheet_name=None)
        except Exception as e:
            log_message(f"Error reading Excel file: {e}")
    else:
        log_message(f"File not found: {selectedFile}")

    
    startDate = start_date_cal.get_date()
    endDate = end_date_cal.get_date()

    if not isinstance(startDate, date):
        startDate = datetime.strptime(startDate, '%m/%d/%y').date()
    if not isinstance(endDate, date):
        endDate = datetime.strptime(endDate, '%m/%d/%y').date()
    
    if retrieve_calender_var.get():
        log_message("Updating staff calenders:")  
        selectedFile = retrieve_calenders(selectedFile, output_text, startDate, endDate)
        print(selectedFile)
        
    solutionDf, capacityUsage, goal_comparison_df, scheduleText = makeSchedule(startDate, endDate, selectedFile, output_text, check_calender=check_calender_var.get(), constant_goal_weight=slider.get(), want_ics=False) #create_ICS.get())

    current_time = datetime.now().strftime("%d%m%H%M")
    if check_calender_var.get(): #If availability taken into account
        suggested_filename = f"Schedule {startDate} to {endDate} with availability - {current_time}.xlsx"
    else:
        suggested_filename = f"Schedule {startDate} to {endDate} without availability - {current_time}.xlsx"
        
    f = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                     initialfile=suggested_filename,  
                                     filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    
    if not f:  # If the user cancels the save dialog, f will be an empty string
        return
    
    # Save the Excel file
    with pd.ExcelWriter(f) as writer:
        solutionDf.to_excel(writer, sheet_name='Schedule', index=False)
        capacityUsage.to_excel(writer, sheet_name='Capacity Usage', index=False)
        goal_comparison_df.to_excel(writer, sheet_name='Goal Comparison', index=False)
                   
        #Active formulas:
        workbook = writer.book
        capacity_usage_worksheet = writer.sheets['Capacity Usage']
        row_count = 2
        for row_num in range(len(capacityUsage)):
            capacity_usage_worksheet.write_formula(row_count-1, 2, f"=SUMIFS(Schedule!E:E, Schedule!G:G, A{row_count}, Schedule!F:F, B{row_count})")
            capacity_usage_worksheet.write_formula(row_count-1, 3, f"=SUMIFS(Schedule!E:E, Schedule!G:G, A{row_count}, Schedule!F:F, B{row_count}, Schedule!C:C, \"Curious Case\")")
            capacity_usage_worksheet.write_formula(row_count-1, 4, f"=SUMIFS(Schedule!E:E, Schedule!G:G, A{row_count}, Schedule!F:F, B{row_count}, Schedule!D:D, \"PAPI1\")")
            capacity_usage_worksheet.write_formula(row_count-1, 5, f"=SUMIFS(Schedule!E:E, Schedule!G:G, A{row_count}, Schedule!F:F, B{row_count}, Schedule!D:D, \"ROLEPLAY1\")")
            capacity_usage_worksheet.write_formula(row_count-1, 6, f"=SUMIFS(Schedule!E:E, Schedule!G:G, A{row_count}, Schedule!F:F, B{row_count}, Schedule!D:D, \"CASE1\")")
            capacity_usage_worksheet.write_formula(row_count-1, 7, f"=SUMIFS(Schedule!E:E, Schedule!G:G, A{row_count}, Schedule!F:F, B{row_count}, Schedule!D:D, \"DATACASE\")")               
            capacity_usage_worksheet.write_formula(row_count-1, 9, f"=I{row_count}-C{row_count}")
            row_count += 1
        
        goal_comparison_worksheet = writer.sheets['Goal Comparison']
        row_count = 2
        for program_name in goal_comparison_df['Program']:
            if program_name != "":
                if program_name == 'Curious': 
                    goal_comparison_worksheet.write_formula(row_count - 1, 3, f"=COUNTIFS(Schedule!C:C, \"Curious Case\", Schedule!G:G, 'Goal Comparison'!A{row_count})/2")
                else:
                    goal_comparison_worksheet.write_formula(row_count-1, 3, f"=COUNTIFS(Schedule!C:C, B{row_count}, Schedule!G:G, 'Goal Comparison'!A{row_count})/3")
                goal_comparison_worksheet.write_formula(row_count-1, 4, f"=D{row_count}-C{row_count}")               
                row_count+=1
        goal_comparison_worksheet.write_formula(row_count-1, 5, f"=(SUM(D:D)/SUM(C:C))*100")              
        goal_comparison_worksheet.write_formula(row_count - 1, 6, f"=(SUMIF(B:B, \"Curious\", D:D) / SUMIF(B:B, \"Curious\", C:C)) * 100")
        goal_comparison_worksheet.write_formula(row_count - 1, 7, f"=(SUM(D:D) - SUMIF(B:B, \"Curious\", D:D)) / (SUM(C:C) - SUMIF(B:B, \"Curious\", C:C)) * 100")

    # Open the saved Excel file automatically
    os.startfile(f)

    # Close the GUI
    app.destroy()

# Initialize the customTKinter application
app = ctk.CTk()
app.title("ORMIT Assessment Scheduler")
window_width = 720*1.3
window_height = 560*1.3
app.geometry(f"{window_width}x{window_height}")
app.resizable(False, True)
# Bring the window to the front and make it topmost
app.lift()
app.attributes('-topmost', True)
app.after(1, lambda: app.attributes('-topmost', False))  # Only topmost temporarily to bring it forward

# Center the window on the screen
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
x_coordinate = int((screen_width / 2) - (window_width / 2))
y_coordinate = int((screen_height / 2) - (window_height / 2))
app.geometry(f'{window_width}x{window_height}+{x_coordinate}+{y_coordinate}')
app.resizable(False, False)  # Make the window not adjustable

ctk.set_default_color_theme(os.path.join(base_path, "resources", "customTheme.json"))
Montserrat = tkFont.Font(family="Montserrat", size=9, weight="bold")
ctk.set_appearance_mode("light")

icon_path = os.path.join(base_path, "resources", "logo.ico")
app.iconbitmap(icon_path)

app.rowconfigure(6, weight=1)
app.columnconfigure(0, weight=1)
app.columnconfigure(1, weight=1)

# File entry
file_entry_label = ctk.CTkLabel(app, text='Assessor Info:')
file_entry_label.grid(row=0, column=1, padx=(20,20), pady=(20,0), sticky="w")
file_entry_var = ctk.StringVar()
file_entry = ctk.CTkEntry(app, width=300, textvariable=file_entry_var, placeholder_text="Select an Excel file...")
file_entry.grid(row=1, column=1,sticky="ew", padx=(20, 20), pady=(0, 5))

# Open button
open_button = ctk.CTkButton(app, text="Browse", command=open_file)
open_button.grid(row=1, column=2, pady=(0, 5))

# Start date label and calendar
start_date_label = ctk.CTkLabel(app, text="Start Date:")
start_date_label.grid(row=2, column=1, padx=(20, 20), pady=(5, 0), sticky="w")
start_date_cal = Calendar(app, selectmode='day',showweeknumbers = False, showothermonthdays = False,
                          year=2025, month=1, day=1,  # Set your default date here
                          font=Montserrat,
                          background = "#343A40", selectbackground ="#003366",
                          headersbackground ="#343A40", headersforeground = "#ffffff")

# start_date_cal = Calendar(app, selectmode='day', showweeknumbers=False, showothermonthdays=False,
#                           year=2025, month=1, day=1,  # Set your default date here
#                           font=Montserrat,
#                           background="#F8F9FA", selectbackground="#003366",
#                           headersbackground="#343A40", headersforeground="#FFFFFF")

start_date_cal.grid(row=3, column=1, padx=(20, 20), pady=(0, 5))

# End date label and calendar
end_date_label = ctk.CTkLabel(app, text="End Date:")
end_date_label.grid(row=2, column=2, padx=(20, 20), pady=(5, 0), sticky="w")
end_date_cal = Calendar(app, selectmode='day',showweeknumbers = False, showothermonthdays = False,
                          year=2025, month=12, day=31,  # Set your default date here
                          font=Montserrat,
                          background = "#343A40", selectbackground ="#003366",
                          headersbackground ="#343A40", headersforeground = "#ffffff")
end_date_cal.grid(row=3, column=2, padx=(20,20), pady=(0, 5))

# Checkbox for using staff information already in file 
check_calender_var = ctk.BooleanVar(value=True)  # Default value can be set to True or False
check_calender_checkbox = ctk.CTkCheckBox(app, text="Consider Staff Availability", variable=check_calender_var)
check_calender_checkbox.grid(row=1, column=0, padx=(20, 20), pady=(10, 0), sticky="w")

# Start scheduling button
start_button = ctk.CTkButton(app, text="Start Scheduling", command=start_scheduling_threaded)
start_button.grid(row=11, column=2, pady=20)

# Checkbox for regenerating live calender situation of staff (Outlook), adds +-4 minutes
retrieve_calender_var = ctk.BooleanVar(value=True)  # Default value can be set to True or False
retrieve_calender_checkbox = ctk.CTkCheckBox(app, text="Update Staff Availability (+4 min.)", variable=retrieve_calender_var)
retrieve_calender_checkbox.grid(row=2, column=0, padx=(20, 20), pady=(10, 0), sticky="w")

# Slider
left_label = ctk.CTkLabel(app, text="Less external | Less goals")
left_label.grid(row=8, column=0, padx=(20, 5), pady=(10, 0), sticky="e")
slider = ctk.CTkSlider(
    app, from_=0.001, to=3, number_of_steps=15, command=slider_callback
)
slider.set(1)  # Set default value

slider.grid(row=8, column=1, padx=(5, 5), pady=(10, 0), sticky="ew")
# Value label to display current slider value
value_label = ctk.CTkLabel(app, text="Value: 1")  # Initial value display
value_label.grid(row=9, column=1, padx=(5, 5), pady=(10, 0), sticky="n")
# Right label
right_label = ctk.CTkLabel(app, text="More external | More goals")
right_label.grid(row=8, column=2, padx=(5, 20), pady=(10, 0), sticky="w")


# # Checkbox for creating ICS files for every assessor if succesful
# create_ICS = ctk.BooleanVar(value=True)  # Default value can be set to True or False
# create_ICS_checkbox = ctk.CTkCheckBox(app, text="Create Outlook files", variable=create_ICS)
# create_ICS_checkbox.grid(row=6, column=0, padx=(20, 20), pady=(10, 0), sticky="w")

#Text window
output_text = ctk.CTkTextbox(app, width=800, height=100, state='disabled')
output_text.grid(row=10, column=0, columnspan=3, padx=(20, 20), pady=(10, 20), sticky="ew")

# Run the application
app.mainloop()