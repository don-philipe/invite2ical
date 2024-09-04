# invite2ical

This little python script scans an IMAP mailbox, filters out emails from a certain sender and extracts the `.ics`-attachements from these mails.
The events found in these ics files will be collected an put into a calendar object.
At the end of the process the calendar object will be written into an ics file on disk.
From the next start of the script on the written ics file will be read first and events from ics attachments will be added to that file if not already present.

The created ics file can be made available with the help of a webserver.
Calendar apps can now subscribe to the URL to add the events to their calendar.
