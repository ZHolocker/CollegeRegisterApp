import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl



def enter_data():
    """This function is called when the user clicks the button""" 
    terms = accept_terms_var.get()
    
    if terms == "accept":
        first_name = first_name_entry.get()
        last_name = last_name_entry.get()
        #Nested if statement to check that names are filled out    
        if first_name and last_name:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()
            #course info
            registration_status = registered_check_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemesters_spinbox.get()
            print(first_name)
            print(last_name)
            print(title)
            print(age)
            print(nationality)
            print(registration_status)
            print(numcourses)
            print(numsemesters)
            print(terms)
            print("-------------------------------")
            #filepath needed to go forward slashes instead of backslashes to work
            filepath =  "C:/Users/MrHol/OneDrive/Desktop/PythonGUIProject/project1/data.xlsx"
            
            #if the file does not exist, create it
            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name", "Last Name", "Title", "Age", "Nationality", "Registration Status",
                           "# of Completed Courses", "# of Completed Semesters", "Terms and Conditions"]
                sheet.append(heading)
                workbook.save(filepath) #save new file after creating
            #if the file does exist, append to it    
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([first_name, last_name, title, age, nationality, registration_status,
                              numcourses, numsemesters, terms])
            workbook.save(filepath) #save file after appending

        else:
            tkinter.messagebox.showwarning(title="Warning", message="Please enter your first and last name")
    else:
        tkinter.messagebox.showwarning(title="Warning", message="You must accept the terms and conditions to submit the form")
    

# this is the main window for the application
window = tkinter.Tk() 
window.title("Data Entry Form")

#this is the frame in the window, and packing it to display the window
frame = tkinter.Frame(window)
frame.pack()

#saving user info
user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row = 0, column = 0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row = 0, column = 0)

last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row = 0, column = 1)

#now to put in the boxes that hold the first and last names
first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)


title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=[" ", "Mr.", "Ms.", "Dr."])
title_label.grid(row = 0, column = 2)
title_combobox.grid(row = 1, column = 2)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row = 2, column = 0)
age_spinbox.grid(row = 3, column = 0)

nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame, values=[" ", "Africa", "Asia", "Europe", "North America", 
                                                             "South America", "Oceania", "Antarctica"])
nationality_label.grid(row = 2, column = 1)
nationality_combobox.grid(row = 3, column = 1)

# a quick way to add padding to the all the widgets
for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)


#Saving course info
courses_info_frame = tkinter.LabelFrame(frame, text="Course Information")
courses_info_frame.grid(row = 1, column = 0, sticky="nsew", padx=20, pady=10)
registered_label = tkinter.Label(courses_info_frame, text="Registration Status")
registered_check_var = tkinter.StringVar(value = "Not registered")
registered_checkbutton = tkinter.Checkbutton(courses_info_frame, text="Currently Registered", 
                                             variable=registered_check_var, onvalue="Registered", offvalue="Not registered")

registered_label.grid(row = 0, column = 0)
registered_checkbutton.grid(row = 1, column = 0)



numcourses_label = tkinter.Label(courses_info_frame, text="# of Completed Courses")
numcourses_spinbox = tkinter.Spinbox(courses_info_frame, from_=0, to='infinity')
numcourses_label.grid(row = 0, column = 1)
numcourses_spinbox.grid(row = 1, column = 1)

numsemesters_label = tkinter.Label(courses_info_frame, text="# of Completed Semesters")
numsemesters_spinbox = tkinter.Spinbox(courses_info_frame, from_=0, to='infinity')
numsemesters_label.grid(row = 0, column = 2)
numsemesters_spinbox.grid(row = 1, column = 2)

for widget in courses_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)



#Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms and Conditions")
terms_frame.grid(row = 2, column = 0, sticky="nsew", padx=20, pady=10)

accept_terms_var = tkinter.StringVar(value="reject")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms and conditions", 
                                  variable=accept_terms_var, onvalue="accept", offvalue="reject")
terms_check.grid(row = 0, column = 0)

#Button
button = tkinter.Button(frame, text="Submit", command= enter_data)
button.grid(row = 3, column = 0, sticky="nsew", padx=20, pady=10)



#this project uses openpyxl and os to connect to an excel file
window.mainloop()
