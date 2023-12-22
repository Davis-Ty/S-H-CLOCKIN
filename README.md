# S-H-CLOCKIN
This project contains two main programs that work togeather.

(The key was to output a program that could help a small business while learning how to push/pull data into/from excel in replace of a database.)

The ClockIn App
Code imports several modules, including threading, tkinter, datetime, os, openpyxl, and xlsxwriter.

The loadup() function reads data from an existing workbook named "ShinningHillEmployee.xlsx". The function reads the data for each employee and puts them into a list called data. The function then creates a new workbook named "ShinningHillEmployeeTimeCard" with the current date and week number appended to the end of the file name. The function then creates a worksheet for each day of the week in the new workbook and adds the data from the data list to each worksheet.

The clockIn() function is called when the user clicks a clock-in button on the GUI. The function reads the ID number of the employee from an Entry widget and searches for the employee's ID in the existing workbook. If the ID is found, the function updates the current date in the workbook, and if the employee has not already clocked in, the function records the current time as the clock-in time for the employee.

Overall, this code is functional program for tracking employee time and attendance.


The ADMIN USER VIEW 

Importing necessary modules:

datetime for working with dates and times.
tkinter for creating the GUI.
threading for running functions in separate threads.
os for interacting with the operating system.
openpyxl for working with Excel files.
xlsxwriter for creating Excel files.
Defining the background_pic function:

This function takes an image file and sets it as the background image of the GUI window.
Defining the adminPass function:

This function is called when the user enters their username and password.
It checks if the entered values match the expected values.
If the values are correct, it defines several other functions related to employee time tracking and payroll.
If the values are incorrect, it shows an error message and clears the username and password fields.
Function definitions within the adminPass function:

calculate_total_pay: Calculates the total pay for an employee based on their clock-in, clock-out, lunch-in, lunch-out times, and hourly rate.
is_valid_id: Checks if a specific row in the spreadsheet has a valid employee ID.
payRoll: Generates a payroll report based on the data in the Excel file and saves it to a text file.
exeDatabase: Checks if the employee file exists; if not, creates it.
loadup: Adds a new employee entry to the file with the provided details.
delete: Deletes an employee entry from the file based on the provided ID.
replace: Replaces an existing employee ID with a new ID in the file.
replacePay: Replaces the pay rate for an employee in the file.

