import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from panopto_folders import PanoptoFolders
from panopto_oauth2 import PanoptoOAuth2
import urllib3
import requests
import json
from datetime import datetime
from urllib.parse import quote
from dateutil import parser, rrule
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import schedule
import argparse
import config
import socket

FINAL_A = "01/14/2022"
FINAL_B = "06/24/2022"


# MAILING
# def document_action(self, body, subject):
#     if config.USER is not None and config.PASSWORD is not None:
#         self.send_mail_and_meeting(subject, body)
#     with open(config.LOG_FILE, 'w') as f:
#         f.write(f'{subject}\n{body}\n\n')
#         f.flush()
#
# def send_mail_and_meeting(self, subject, body):
#     email_sender = config.USER
#
#     msg = MIMEMultipart()
#     msg['From'] = email_sender
#     msg['To'] = ", ".join(config.TO_SEND)
#     msg['Subject'] = subject
#
#     msg.attach(MIMEText(f'{body}\n {str(self)}', 'plain'))
#     try:
#         server = smtplib.SMTP('smtp.office365.com', 587)
#         server.ehlo()
#         server.starttls()
#         server.login(config.USER, config.PASSWORD)
#         text = msg.as_string()
#         server.sendmail(email_sender, config.TO_SEND, text)
#         print('email sent')
#         server.quit()
#     except socket.error:
#         print("SMPT server connection error")
#     return True
#
# def delete_all(self):
#     for session_id in self.sessions_id:
#         url = config.BASE_URL + "scheduledRecordings/{0}".format(session_id)
#         print('Calling DELETE {0}'.format(url))
#         read_resp = requests_session.delete(url=url).json()
#         print("DELETE returned:\n" + json.dumps(read_resp, indent=2))


def authorization(requests_session, oauth2):
    # Go through authorization
    access_token = oauth2.get_access_token_authorization_code_grant()
    # Set the token as the header of requests
    requests_session.headers.update({'Authorization': 'Bearer ' + access_token})


def search(course_id, year, semester):
    if semester == 'א':
        semester = "Semester 1"
    elif semester == 'ב':
        semester = "Semester 2"
    else:
        semester = "Semester 3"

    results = folders.search_folders(rf'{course_id}')
    id = None
    for result in results:
        if str(year) in result['Name']:
            if result['ParentFolder']['Name'] == f'{year} -> {semester}' or \
                    result['ParentFolder']['Name'] == f'{year} -> Semester 1 or 2' or \
                    result['ParentFolder']['Name'] == f'{year} -> Semesters 1 or 2' or \
                    result['ParentFolder']['Name'] == f'{year} -> Non-shnaton':
                return result['Id']
    return id


def update_client():
    global client
    client = gspread.authorize(creds)


# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
def schedule_all():
    sheet = client.open("Recordings").sheet1
    data = pd.DataFrame(sheet.get_all_records())
    # sheet.delete_rows(2, data.shape[0] + 1)
    if data.empty:
        return
    data.rename(columns=config.COLUMN_NAMES, inplace=True)
    for tuple_row in data.iterrows():
        time_stamp, course_number, semester, hall, date, time_beginning, time_end, which_days = tuple_row[1]
        schedule_request(time_stamp, course_number, semester, hall, date, time_beginning, time_end,
                         which_days)
    sheet.delete_rows(2, data.shape[0] + 1)
    return


def schedule_request(time_stamp, course_number, semester, hall, date, time_beginning, time_end, which_days):
    recorder = config.SERVERS[hall]
    url = config.BASE_URL + "remoteRecorders/search?searchQuery={0}".format(quote(recorder))
    print('Calling GET {0}'.format(url))
    resp = requests_session.get(url=url).json()
    recorder = [rr for rr in resp['Results'] if rr['Name'] == recorder]
    if 'Results' not in resp or len(recorder) != 1:
        print("Recorder not found:\n{0}".format(resp))
        return None
    recorder = recorder[0]
    folder_id = search(course_number, "2020-21", semester)
    # FORMAT - MM/DD/YYYY
    # date = datetime.strptime(date, "%d/%m/%Y").strftime("%m-%d-%Y")  Depends on Date format
    start_date_str = f'{date} {time_beginning}'
    end_date_str = f'{date} {time_end}'

    datetime_start = parser.parse(start_date_str, dayfirst=True)
    datetime_end = parser.parse(end_date_str, dayfirst=True)
    start_time = config.ISRAEL.localize(datetime_start)
    end_time = config.ISRAEL.localize(datetime_end)
    schedule(recorder, start_time, end_time, which_days, folder_id, course_number, semester, time_beginning, time_end)


def schedule(recorder_server, start_date: datetime, end_date: datetime, which_days, folder_id, course_number, semester,
             start, end):
    if which_days:
        # pass
        if semester == 'א':
            end_start_A = f'{FINAL_A} {start}'
            end_end_A = f'{FINAL_A} {end}'
            datetime_start = parser.parse(end_start_A, dayfirst=True)
            datetime_end = parser.parse(end_end_A, dayfirst=True)
            start_time = config.ISRAEL.localize(datetime_start)
            end_time = config.ISRAEL.localize(datetime_end)
            start_dates = [start for start in rrule.rrule(rrule.WEEKLY, dtstart=start_date, until=start_time)]
            end_dates = [end for end in rrule.rrule(rrule.WEEKLY, dtstart=end_date, until=end_time)]
        # else:
        #     start_dates = [start for start in rrule.rrule(rrule.WEEKLY, dtstart=start_date, until=FINAL_B)]
        #     end_dates = [end for end in rrule.rrule(rrule.WEEKLY, dtstart=end_date, until=FINAL_B)]
    else:
        start_dates = [start_date]
        end_dates = [end_date]

    for start_date, end_date in zip(start_dates, end_dates):
        sr = {
            "Name": str(course_number)+" "+str(start_date)[:-9],
            "Description": "",
            "StartTime": start_date.isoformat(),
            "EndTime": end_date.isoformat(),
            "FolderId": folder_id,
            'Recorders': [
                {
                    'RemoteRecorderId': recorder_server['Id'],
                    'SuppressPrimary': False,
                    'SuppressSecondary': True,
                }
            ],
            'IsBroadcast': True,
        }
        url = config.BASE_URL + "scheduledRecordings?resolveConflicts=false"
        print('Calling POST {0}'.format(url))
        create_resp = requests_session.post(url=url, json=sr).json()
        print("POST returned:\n" + json.dumps(create_resp, indent=2))
        if 'Id' not in create_resp:
            print('CANT SCHEDULE')
            return
    print("SUCCESS")
    return


def main():
    global folders, client, creds, requests_session
    # use creds to create a client to interact with the Google Drive API
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(config.GOOGLE_JSON, scope)
    client = gspread.authorize(creds)
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    # Use requests module's Session object in this example.
    # ref. https://2.python-requests.org/en/master/user/advanced/#session-objects
    requests_session = requests.Session()

    # Load OAuth2 logic
    oauth2 = PanoptoOAuth2(config.PANOPTO_SERVER_NAME, config.PANOPTO_CLIEND_ID, config.PANOPTO_SECRET, False)
    authorization(requests_session, oauth2)
    # Load Folders API logic
    folders = PanoptoFolders(config.PANOPTO_SERVER_NAME, False, oauth2)

    # parse_argument() אילן הוסיף אבל לא צריך
    schedule_all()
    # schedule.every().minute.do(schedule_all)
    # schedule.every(1).hours.do(update_client)
    # schedule.every(1).hours.do(authorization, requests_session, oauth2)
    # while True:
    #     schedule.run_pending()


if __name__ == '__main__':
    main()