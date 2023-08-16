# Reservation
import win32com.client
import datetime
import calendar

def check_and_reserve(room_email, start_datetime, duration_minutes, subject):
    oOutlook = win32com.client.Dispatch("Outlook.Application")
    
    appt = oOutlook.CreateItem(1)
    appt.Start = start_datetime
    appt.Duration = duration_minutes
    appt.MeetingStatus = 1
    appt.Subject = subject
    
    myRecipient = appt.Recipients.Add(room_email)
    myRecipient.Resolve()
    
    if myRecipient.FreeBusy(start_datetime, duration_minutes, True)[0] == 0:
        appt.Save()
        appt.Send()
        print("Reservation sent successfully.")
    else:
        print("Room not available.")

room_email = "<mail id of meeting room>"
duration_minutes = 60
subject = "Meeting Reservation"

# Specify the year and month
year = 2023
month = 8

# Get the number of days in the specified month
days_in_month = calendar.monthrange(year, month)[1]

# Check availability and make reservations for every Tuesday
for day in range(1, days_in_month + 1):
    date = datetime.datetime(year, month, day)
    if date.weekday() == 1:  # Tuesday is represented by 1
        check_and_reserve(room_email, date, duration_minutes, subject)
