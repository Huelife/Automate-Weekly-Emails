# Automate-Weekly-Emails
Automate sending weekly pdf meeting logs in VBA, using Excel, Adobe, and Outlook.

Automating through Task Scheduler:

1) Create Basic Task:<br>
   A) Name: Weekly Logs Report<br>
   B) Description: PDF to Outlook reoccuring pre-meeting weekly logs<br>
2) Trigger:<br>
   A) Weekly (Check this one for weekly reports)<br>
   B) Start: 10/20/2023 at 1:00:00 PM<br>
   C) Recur every: 1 (Put 1 for every week) weeks on:<br>
   D) Thursday (Check this one for sending reports out on Thursdays)<br>
3) Action:<br>
   A) Start a program (Check this one)<br>
   B) Program/script:<br>
      (email_logs.bat)<br>
   C) Start in (optional):<br>
      (C:\Users\to\.bat\file\location\)<br>
5) Finish<br>
