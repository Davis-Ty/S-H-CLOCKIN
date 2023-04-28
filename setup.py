from cx_Freeze import setup, Executable

executables=[Executable("OneDrive/Desktop/S-H-CLOCKIN/Clock_In_App.py"),Executable("OneDrive/Desktop/S-H-CLOCKIN/Admin_View.py")]


setup(
    name="Better Dayz",
    version="1.0",
    description="This code seems to be a time tracking application that allows employees to clock in and out, take breaks, and track their time. It uses the tkinter library for the graphical user interface and the openpyxl library for reading and writing to an Excel file. It also uses threading to display output messages for a short amount of time and then hide them. The application appears to be using a spreadsheet to store the employee information and time data. There are four functions for clocking in, clocking out, taking lunch, and ending lunch. Each function checks if the user's ID is valid and if they meet the necessary requirements to clock in, clock out, or take a lunch break. The application also saves the data to an Excel file.",
    executables=executables,
        options={
        'build_exe': {
            'packages': [],
            'include_files': ['OneDrive/Desktop/S-H-CLOCKIN/ShinningHillPic.png']
        }
    }
)


