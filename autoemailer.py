#!/usr/bin/env python

import datetime
import time
import urllib.request
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import xmltodict

import secret


def get_timeslot_info(date):
    timeslot_list = []
    url_string = (
        'https://awh-vogel.sona-systems.com/services/SonaAPI.svc/GetStudyScheduleList?'
        f'username={secret.sona_username}&password={secret.sona_password}'
        f'&location_id=-2&start_date={date}&end_date={date}&lab_only=0'
    )

    with urllib.request.urlopen(url_string) as xml_file:
        xml_data = xml_file.read()
    response = xmltodict.parse(xml_data)
    study_info = response["GetStudyScheduleListResponse"]["GetStudyScheduleListResult"]["a:Result"]

    for key in study_info:
        if key == "a:APIStudySchedule":
            day = study_info["a:APIStudySchedule"]
            if type(day) is list:
                for timeslot in day:
                    timeslot_dict = {}
                    timeslot_dict["experiment_ID"] = timeslot["a:experiment_id"]
                    timeslot_dict["study_name"] = timeslot["a:study_name"]
                    timeslot_dict["location"] = timeslot["a:location"]
                    timeslot_dict["ID"] = timeslot["a:timeslot_id"]
                    timeslot_dict["researcher_ID"] = timeslot["a:researcher_id"]
                    timeslot_dict["datetime"] = datetime.datetime.strptime(
                        timeslot["a:timeslot_date"], '%Y-%m-%dT%H:%M:%S'
                    )
                    timeslot_list.append(timeslot_dict)
            else:
                timeslot = day
                timeslot_dict = {}
                timeslot_dict["experiment_ID"] = timeslot["a:experiment_id"]
                timeslot_dict["study_name"] = timeslot["a:study_name"]
                timeslot_dict["location"] = timeslot["a:location"]
                timeslot_dict["ID"] = timeslot["a:timeslot_id"]
                timeslot_dict["researcher_ID"] = timeslot["a:researcher_id"]
                timeslot_dict["datetime"] = datetime.datetime.strptime(
                    timeslot["a:timeslot_date"], '%Y-%m-%dT%H:%M:%S'
                )
                timeslot_list.append(timeslot_dict)
        else:
            continue

    return timeslot_list


def get_participants(timeslot_dict):
    participants = []
    url_string = (
        'https://awh-vogel.sona-systems.com/services/SonaAPI.svc/GetSignUpsForTimeslot?'
        f'username={secret.sona_username}&password={secret.sona_password}&timeslot_id={timeslot_dict["ID"]}'
     )
    with urllib.request.urlopen(url_string) as xml_file:
        xml_data = xml_file.read()
    response = xmltodict.parse(xml_data)
    timeslot_info = response["GetSignUpsForTimeslotResponse"]["GetSignUpsForTimeslotResult"]["a:Result"]

    for key in timeslot_info:
        if key == "a:APISignUp":
            timeslot = timeslot_info["a:APISignUp"]
            if type(timeslot) is list:
                for participant in timeslot:
                    participant_dict = {}
                    participant_dict['user_ID'] = participant["a:user_id"]
                    participant_dict['first_name'] = participant["a:first_name"]
                    participant_dict['last_name'] = participant["a:last_name"]
                    participant_dict['email'] = participant["a:email"]
                    participant_dict['study_name'] = timeslot_dict['study_name']
                    participant_dict['location'] = timeslot_dict['location']
                    participant_dict['datetime'] = timeslot_dict['datetime']
                    participant_dict["experiment_ID"] = timeslot_dict["experiment_ID"]
                    participant_dict["researcher_email"] = get_researcher_email(
                        participant_dict["experiment_ID"], timeslot_dict["researcher_ID"]
                    )
                    participants.append(participant_dict)
            else:
                participant = timeslot
                participant_dict = {}
                participant_dict['user_ID'] = participant["a:user_id"]
                participant_dict['first_name'] = participant["a:first_name"]
                participant_dict['last_name'] = participant["a:last_name"]
                participant_dict['email'] = participant["a:email"]
                participant_dict['study_name'] = timeslot_dict['study_name']
                participant_dict['location'] = timeslot_dict['location']
                participant_dict['datetime'] = timeslot_dict['datetime']
                participant_dict["experiment_ID"] = timeslot_dict["experiment_ID"]
                participant_dict["researcher_email"] = get_researcher_email(
                    participant_dict["experiment_ID"], timeslot_dict["researcher_ID"]
                )
                participants.append(participant_dict)

    good_participants = []

    for participant in participants:
        if is_invalid_account(participant):
            send_invalid_participant_email(participant)
        else:
            good_participants.append(participant)

    return good_participants


def is_invalid_account(participant):
    url_string = (
        'https://awh-vogel.sona-systems.com/services/SonaAPI.svc/GetPersonInfoByUserID?'
        f'username={secret.sona_username}&password={secret.sona_password}&user_id={participant["user_ID"]}'
    )
    with urllib.request.urlopen(url_string) as xml_file:
        xml_data = xml_file.read()
    participant_info = xmltodict.parse(xml_data)
    unexcused = (
        participant_info
        ['GetPersonInfoByUserIDResponse']
        ['GetPersonInfoByUserIDResult']
        ['a:Result']
        ['a:APIPerson']
        ['a:noshow_count']
    )
    return int(unexcused) >= 2


def send_invalid_participant_email(participant):
    from_email = secret.gmail_address
    to_email = participant['researcher_email']
    cc_email = secret.admin_address + ',' + secret.gmail_address

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(secret.gmail_name, secret.gmail_password)

    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = to_email
    message["Cc"] = cc_email
    message["Subject"] = 'Warning: One of your subjects has 2 Unexcused Absences'

    body = (
        f'Your participant {participant["first_name"]} {participant["last_name"]} '
        f'is signed up for the study "{participant["study_name"]}" tomorrow.\n'
        'The sona records show that this participant has 2 or more unexcused absences.\n\n'
        'If you are an RA, please forward this email to a grad student or post doc.'
    )

    message.attach(MIMEText(body))
    email_text = message.as_string()

    server.sendmail(from_email, [to_email] + cc_email.split(','), email_text)


def get_researcher_email(exp_ID, first_researcher_ID):
    ignored_researchers = [267, 2]  # ignore Ariana and Brendan

    if first_researcher_ID == "0":  # I forget what this is checking for...
        url_string = (
            'https://awh-vogel.sona-systems.com/services/SonaAPI.svc/GetResearcherIDByExperimentID?'
            f'username={secret.sona_username}&password={secret.sona_password}&experiment_id={exp_ID}'
        )

        with urllib.request.urlopen(url_string) as xml_file:
            xml_data = xml_file.read()
        experiment_info = xmltodict.parse(xml_data)

        researcher_ID = (
            experiment_info
            ['GetResearcherIDByExperimentIDResponse']
            ['GetResearcherIDByExperimentIDResult']
            ['a:Result']
            ['b:int']
        )

        if not isinstance(researcher_ID, str):
            i = 0
            while researcher_ID[i] in ignored_researchers:
                i += 1
                if i >= len(researcher_ID):  # If it's only ignored researchers, just use the first one
                    i = 0
                    break
            researcher_ID = researcher_ID[i]
    else:
        researcher_ID = first_researcher_ID

    url_string = (
        'https://awh-vogel.sona-systems.com/services/SonaAPI.svc/GetPersonInfoByPersonID?'
        f'username={secret.sona_username}&password={secret.sona_password}&person_id={researcher_ID}'
     )

    with urllib.request.urlopen(url_string) as xml_file:
        xml_data = xml_file.read()
    researcher_info = xmltodict.parse(xml_data)

    researcher_email = (
        researcher_info
        ['GetPersonInfoByPersonIDResponse']
        ['GetPersonInfoByPersonIDResult']
        ['a:Result']
        ['a:APIPerson']
        ['a:alt_email']
    )

    return researcher_email


def send_emails(participant_list):
    retry_list = []

    from_email = secret.gmail_address

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(secret.gmail_name, secret.gmail_password)

    for participant in participant_list:
        to_email = participant['email']

        message = MIMEMultipart()
        message["From"] = from_email
        message["To"] = to_email
        message["reply-to"] = participant['researcher_email']
        message["Subject"] = (
            'Awh Vogel Lab - Reminder for your study at '
            f'{datetime.datetime.strftime(participant["datetime"],"%I:%M %p")} tomorrow'
        )

        body = (
            f'Hi {participant["first_name"]},\n\n'
            'This is a reminder that you are signed up for a study tomorrow.\n\n'
            f'Study Name: {participant["study_name"]}\n'
            f'Researcher Email: {participant["researcher_email"]}\n'
            f'Location: {participant["location"]}\n'
            f'Time: {datetime.datetime.strftime(participant["datetime"],"%A, %B %d, %Y %I:%M %p")}\n\n'
            'Here is a map to our building: https://maps.uchicago.edu/?location=Biopsychological+Research+Building\n'
            'Please wait in the lobby on the bench by the elevator.\n'
            'If you arrive early, please wait until your scheduled time.\n'
            'If you have a problem getting to the lab on time or your experimenter has not arrived 5 minutes after '
            'your scheduled time, please call the lab at (773) 795-4784.\n\n'
            'If you wish to cancel you have until 5pm today. If you need to cancel after 5pm, please reply to '
            'this email (make sure it is going to your researcher and not the do not reply email address).\n'
            'If it is not an emergency it will be counted as an unexcused absence.\n\n'
            'Please note: After 2 unexcused no-shows you will be removed from the system and '
            'will be unable to participate in studies.\n'
            'We will also remove you from our system if we are unable to use your data '
            '(due to excessive blinks/eye movements or not trying to do the task) in 3 sessions.\n'
            'You will receive a warning if 2 of your sessions cannot be used.\n\n'
            'For more information please log on to our site: https://awh-vogel.sona-systems.com\n\n'
            'See you tomorrow!\n\n'
            'Thanks,\n'
            'Awh Vogel Lab'
        )

        message.attach(MIMEText(body))
        email_text = message.as_string()

        try:
            server.sendmail(from_email, to_email, email_text)
        except Exception as e:
            retry_list.append(participant)
            continue

    return retry_list


def send_error_alert(e):
    from_email = secret.gmail_address
    to_email = secret.admin_address
    cc_email = secret.gmail_address

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(secret.gmail_name, secret.gmail_password)

    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = to_email
    message["Cc"] = cc_email
    message["Subject"] = 'ERROR: An error has occured with the reminder email system.'

    body = (
        'There has been an error while trying to email participants.\n\n'
        'Here is the error text:\n\n'
        f'{repr(e)}'
    )

    message.attach(MIMEText(body))
    email_text = message.as_string()

    server.sendmail(from_email, [to_email] + [cc_email], email_text)


def send_success_email():
    from_email = secret.gmail_address
    to_email = secret.gmail_address

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(secret.gmail_name, secret.gmail_password)

    today = datetime.datetime.now().strftime('%x')

    message = MIMEMultipart()
    message["From"] = from_email
    message["To"] = to_email
    message["Subject"] = f'The Reminder Email System has finished executing successfully ({today}).'

    body = (
        'No errors occured while sending emails.\n'
        'If no emails were sent, check that there are sign ups for tomorrow.'
    )

    message.attach(MIMEText(body))
    email_text = message.as_string()

    server.sendmail(from_email, to_email, email_text)


def main():
    tomorrow_date = datetime.date.today() + datetime.timedelta(days=1)
    tomorrow_string = tomorrow_date.strftime('%Y-%m-%d')

    timeslot_list = get_timeslot_info(tomorrow_string)

    participant_list = []
    for timeslot in timeslot_list:
        participants = get_participants(timeslot)
        for participant in participants:
            participant_list.append(participant)

    retry_list = send_emails(participant_list)

    tmp = 0

    while retry_list:
        tmp += 1
        time.sleep(3600)
        send_emails(retry_list)
        if tmp > 3:
            raise EnvironmentError('Some subjects could not be reached.')
            break


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        send_error_alert(e)
        print('There has been an error while emailing the participants.\n\n')
        raise
    else:
        send_success_email()
