# gmail-to-sheets

Simple Google Apps Script that extracts key fields from unread emails and stores it in Google Sheets. The emails are left unread.

To use this script:

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Copy and paste this code into the script editor
4. Save the project
5. Refresh your Google Sheet
6. You'll see a new menu item called "Gmail Processing"
7. Click "Process Unread Emails" to run the script

The script will create a new row in your sheet for each unread email, with the following columns:
- Column A: Date
- Column B: Sender
- Column C: Subject
- Column D: Email body

Note: By default, the script will keep emails unread. 