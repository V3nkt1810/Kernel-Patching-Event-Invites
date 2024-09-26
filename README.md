This repository contains a Python script that automates the process of sending calendar invites for kernel patching events on multiple servers using Microsoft Teams. The script reads server details from an Excel sheet, filters them based on patching dates, and sends out meeting invitations with the necessary details to all required and optional recipients.

Script: Kernel_Patch_Event_Invite.py
Purpose:
The script is designed to:

Send automated calendar invites for kernel patching events to relevant stakeholders.
Pull server details (including patching schedules) from an Excel sheet.
Create and send calendar invites in .ics format via email for each patching date and environment.
Customize each email with server information, including the date, environment, and the meeting link for Microsoft Teams.
Features:
ICS Calendar Invite Creation:

The script generates .ics calendar invites for each scheduled patching event.
Dynamic Email Generation:

Emails are personalized with the meeting date, server environment details, and the list of attendees.
Excel Integration:

Reads server data from an Excel file, including columns for required and optional invitees.
Error Logging:

Detailed logging using Python’s logging module helps track the status of email invites and any potential errors during execution.
Microsoft Teams Meeting Link:

Automatically inserts the provided Teams meeting link into the email body and the calendar invite.
Usage:
Clone the repository:

bash
Copy code
git clone https://github.com/your-repo/Kernel_Patch_Event_Invite.git
cd Kernel_Patch_Event_Invite
Prepare the Excel file:

Ensure your Excel file has the following columns:
Date (Patch date for each server)
Required (Required attendees)
Optional (Optional attendees)
env (Environment name for servers)
Configure the script:

Update the following variables in the script:
subject: The subject of the meeting invite.
server_details_file: Path to your Excel file.
sheet_name: The sheet name in the Excel file containing server details.
meet_link: The Microsoft Teams meeting link.
