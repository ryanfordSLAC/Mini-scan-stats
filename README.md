# Mini-scan-stats
Crunch stats for scans at the various mini-scanner boxes

Use the task scheduler to run the batch file on a recurring (Daily) basis.

In order to crunch the box data, set up a Power Automate Cloud Flow at Office.com to strip attachments from the daily scanner box emails and save them into your one drive.  Then, point the python file to that location.
