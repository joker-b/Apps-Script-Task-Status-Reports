Apps-Script-Task-Status-Reports
===============================

Send periodic email reports about ongoing tasks in your Google Tasks lists. A one-file tool for Google Apps Script.

(Note: Google Tasks is an "Advanced API" usage, you'll need to explicitly turn it on via the API Control panel -- instructions should appear the first time you run the script)

This assumes that you are keeping your task list in Google Calendar/Google Tasks (I also use the "Jorte" Android calendar, which integrates to Google Tasks)

Code the file "Code.gs" into your own task status project in Google Drive.

Edit the takslist names and email adresses found at the top to your own names and addresses.

Test by running one of the top functions for weekly reports (e.g. send_all())

For ongoing use, set a trigger to fire-off reports on a weekly basis. I set mine to run on Friday afternoons.

The script now also offers daily reporting -- I send myself a daily reminder every morning

