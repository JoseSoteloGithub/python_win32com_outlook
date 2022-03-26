# python_win32com_outlook

Run main.py

What this does:
Using the win32com module (pip install pywin32), iterates through the inbox of Outlook and loads the data into an Excel workbook.

The win32com module is the closest module that I found that is similar to VBA associated with the Microsoft Office Suite

To Do Feature:
To use the To Do feature, edit the body of an individual email and type 'TO DO: Some Task Needed' separated by a new line for multiple items, 'TO DO: Another task needed'

To Edit an email OutlookEmailItem>Move>Edit Message

This will export the email information included it's added To Do tasks into separate rows
