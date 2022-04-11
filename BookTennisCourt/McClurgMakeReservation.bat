@echo off
SET court_number=%1
SET list_appointment_military_hour=%2

:: To be safe, it would be better to specify the python.exe path, in case there are two or more version of python installed on the machine
:: and both pathes are listed in environment variable PATH. 
:: python C:\Projects\GitHub\UtilityRepository\BookTennisCourt\BookTennisCourt.py %court_number% %list_appointment_military_hour%

python "C:\Users\Me\Projects\GitHub\Utility\BookTennisCourt\McClurgMakeReservation.py" %court_number% %list_appointment_military_hour%