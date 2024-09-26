import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime
import logging
import pytz
import re
import numpy as np
 
logging.basicConfig(filename='event_invite.log', level=logging.DEBUG, format='%(asctime)s:%(levelname)s:%(message)s')
 
def create_ics_content(subject, start_time, end_time, location, description, attendees):
    logging.debug("Creating ICS content")
    attendee_list = "\n".join([f"ATTENDEE;RSVP=TRUE:mailto:{email}" for email in attendees])
    return f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Your Company//Your Product//EN
METHOD:REQUEST
BEGIN:VEVENT
UID:{datetime.now().strftime('%Y%m%dT%H%M%S')}@yourcompany.com
DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%S')}
DTSTART:{start_time.strftime('%Y%m%dT%H%M%S')}
DTEND:{end_time.strftime('%Y%m%dT%H%M%S')}
SUMMARY:{subject}
LOCATION:{location}
DESCRIPTION:{description}
TRANSP:TRANSPARENT
STATUS:FREE
{attendee_list}
END:VEVENT
END:VCALENDAR"""
 
def send_email(subject, recipients, df, date, meet_link):
    logging.debug("Preparing to send email")
    sender_email = ""
    smtp_server = ""
    smtp_port = 25
 
    est = pytz.timezone('US/Eastern')
    date = pd.to_datetime(date)
    start_time = est.localize(date.replace(hour=0, minute=0, second=0))
    end_time = est.localize(date.replace(hour=7, minute=0, second=0))
 
    df_without_date = df.drop(columns=['Date', 'Required', 'Optional'])
    html_table = df_without_date.to_html(index=False, escape=False)
    html_table = html_table.replace('<thead>', '<thead style="color:lightviolet;">')
 
    html_description = f"""
    <html>
    <head>
        <style>
            .important {{ font-weight: bold; color: red; }}
        </style>
    </head>
    <body>
        <p>Hi All,</p>
        <p>This meeting is scheduled on  {date.strftime("%d/%B/%Y")} to 7 AM EST {date.strftime("%d/%B/%Y")}.</p>
        <p>If this is inconvenient and will cause an impact to the business, please respond with an approval ticket so that it can be rescheduled.\n</p>
        {html_table}
        <p class="important"><b>Location:</b> <a href="{meet_link}">Microsoft Teams Meeting</a></p>
        <p>----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------</p>
        <p>Organizer: Venkatesh S</p>
        <p>Email: </p>
        <p>Note: This is an Automated email</p>
        <p><b>Start Time:</b> {start_time.strftime("%Y-%m-%d %I:%M %p")} EST</p>
        <p><b>End Time:</b> {end_time.strftime("%Y-%m-%d %I:%M %p")} EST</p>
    </body>
    </html>
    """
 
    all_recipients = list(set(recipients['To'].split(',')) | set(recipients['CC'].split(',')))
    ics_content = create_ics_content(subject, start_time, end_time, meet_link, html_description, recipients['To'].split(','))
 
    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipients['To']
    msg['CC'] = recipients['CC']
 
    msg.attach(MIMEText(html_description, 'html'))
 
    ical_attachment = MIMEText(ics_content, 'calendar;method=REQUEST')
    ical_attachment.add_header('Content-Disposition', 'attachment; filename="invite.ics"')
    msg.attach(ical_attachment)
 
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.sendmail(sender_email, all_recipients, msg.as_string())
            logging.info("Event invite sent successfully to %s", ', '.join(all_recipients))
    except Exception as e:
        logging.error("Failed to send invite. Error: %s", str(e))
 
def send_teams_event(subject, server_details_file, sheet_name, meet_link):
    logging.debug("Reading server details from Excel file")
    try:
        df = pd.read_excel(server_details_file, sheet_name=sheet_name, engine='openpyxl')
        df['Date'] = pd.to_datetime(df['Date'])
 
        unique_dates = df['Date'].dt.date.unique()
        for date in unique_dates:
            filtered_df = df[df['Date'].dt.date == date]
            unique_envs = '/'.join(sorted(set(filtered_df['env'].dropna())))
            date_str = date.strftime("%d-%m-%Y")
            updated_subject = f"Kernel-Patching for {unique_envs} Servers on {date_str}"
 
            required_recipients = filtered_df['Required'].fillna('').unique()
            optional_recipients = filtered_df['Optional'].fillna('').unique()
 
            recipients = {
                'To': ','.join(set(required_recipients)),
                'CC': ','.join(set(optional_recipients))
            }
 
            logging.debug(f"Sending email for date {date_str} with recipients: {recipients}")
            send_email(updated_subject, recipients, filtered_df, date, meet_link)
    except Exception as e:
        logging.error("Error processing the Excel file or sending invites: %s", str(e))
 
if __name__ == "__main__":
    logging.debug("Starting the script")
    subject = ""
    server_details_file = "data.xlsx"
    sheet_name = "SERVER DATA"
    meet_link = ""
    send_teams_event(subject, server_details_file, sheet_name, meet_link)
