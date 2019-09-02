import imaplib
import email
from EmailMessage import emailMessage
import re
from datetime import datetime
import time
import json

# Variables to store account details
HOST = 'imap.gmail.com'
EMAIL = 'attendance@softinfo.in'
PASSWORD = 'att@softinfo'


# Function to read and store all unread emails having the subject "Attendance"
def read_mails(mail):
    # Select the inbox
    mail.select('inbox')

    # Assign a unique ID to all the emails matching the filter (unseen and subject = "attendance")
    typ, data = mail.uid('search', None, '(UNSEEN)')  #The correct replacement for filter : (UNSEEN) SUBJECT Attendance
    email_messages = []
    sender_names = []
    sender_ids = []

    if len(data[0]) > 0: #FIX: condition for no unread emails.
        # Unique IDs of all the mails
        mail_uids = data[0].split()

        # Iterate through each UID and fetch the mail
        for mail_uid in mail_uids:
            typ, data = mail.uid('fetch', mail_uid, '(RFC822)')
            raw_email = data[0][1].decode('utf-8')

            # After decoding the mail, store it as a string
            email_message_full = email.message_from_string(raw_email)

            # The class stores the from, to, subject, and body part of the email, we store the body and sender 
            emailObject = emailMessage(email_message_full)
            email_message = emailObject.body
            sender_name = emailObject.name
           # sender_id = emailObject.emailID

            email_messages.append(email_message)
            sender_names.append(sender_name)
           # sender_ids.append(sender_id)
    else:
        print('No new emails')

    return email_messages, sender_names


# Function to retrieve uid, date, in time, out time, and remarks from the email messages provided
def get_details(email_messages, sender_names):
    attendances = []
    for email_message, sender_name in zip(email_messages, sender_names):
        # Put the name of the sender in the dictionary
        attendance = {}
        attendance["Name"] = sender_name.strip(' ').strip()
        
        # Find UID, Date, In time, Out time, Remarks
        uid_index = email_message.upper().find('UID:')
        try:
            date_index = email_message.upper().find('VISIT DATE:')
        except:
            date_index = email_message.upper().find('VISITDATE:')
        try:
            in_index = email_message.upper().find('IN TIME:')
        except:
            in_index = email_message.upper().find('INTIME:')
        try:
            out_index = email_message.upper().find('OUT TIME:')
        except:
            out_index = email_message.upper().find('OUTTIME:')

        remarks_index = email_message.upper().find('REMARKS:')

        if uid_index != -1:
            endline_index = email_message.find('\n', uid_index) # search for \n starting from uid index
            attendance['UID'] = email_message[uid_index + len('UID:') : endline_index].strip()
            # remove anything that is not a number
            attendance['UID'] = re.sub(r'[^0-9]+', '', attendance['UID'])
        else:
            print('Error: UID not found')
            attendance['UID'] = '000'

        if date_index != -1:
            endline_index = email_message.find('\n', date_index) # search for \n starting from date index
            attendance['Date'] = email_message[date_index + len('DATE:') : endline_index].strip()
            #remove anything that is not a number or - or /
            attendance['Date'] = re.sub(r"[^0-9//-]+",'',attendance['Date'])
        else:
            print('Error: Date not found')
            attendance['Date'] = '00/00/00'

        if in_index != -1:
            endline_index = email_message.find('\n', in_index) # search for \n starting from in index
            attendance['In Time'] = email_message[in_index + len('IN TIME:') : endline_index].strip()
            #remove anything that is not a number or colon
            attendance['In Time'] = re.sub(r"[^0-9pmPMamAM:]+",'',attendance['In Time'])
        else:
            print('Error: In time not found')
            attendance['In Time'] = '00:00'

        if out_index != -1:
            endline_index = email_message.find('\n', out_index) # search for \n starting from out index
            attendance['Out Time'] = email_message[out_index + len('OUT TIME:') : endline_index].strip()
            #remove anything that is not a number or colon
            attendance['Out Time'] = re.sub(r"[^0-9pmPMamAM:]+",'',attendance['Out Time'])
        else:
            print('Error: Out time not found')
            attendance['Out Time'] = '00:00'
        
        if remarks_index != -1:
            attendance['Remarks'] = email_message[remarks_index + len('REMARKS:') : ].strip()
        else:
            print('Error: Remarks not found')
            attendance['Remarks'] = ''
        
        attendances.append(attendance)

    return attendances


# function to adjust date and time to uniform form
def rectify_attendances(attendances):
    for attendance in attendances:
        # Create time objects
        in_time = None
        out_time = None
        # Formats accepted: 1:05 pm, 1:05pm, 13:05 
        # strptime creates an object from the provided format fmt
        # strftime returns the object as a string as the provided format
        for fmt in ("%I:%M%p", "%H:%M", "%H", "%I%p", "%H%M", "%I%M%p"):
            try:
                in_time = datetime.strptime(attendance['In Time'].replace('.', '').replace(' ', ''), fmt)
            except ValueError:
                pass

        if not in_time:
            print('Invalid format for In Time: ', attendance['In Time'])
            in_time = datetime.strptime('00:00', '%H:%M')


        for fmt in ("%I:%M%p", "%H:%M", "%H", "%I%p", "%H%M", "%I%M%p"):
            try:
                out_time = datetime.strptime(attendance['Out Time'].replace('.', '').replace(' ', ''), fmt)
            except ValueError:
                pass
        
        if not out_time:
           print('Invalid format for Out Time:', attendance['Out Time'])
           out_time = datetime.strptime('00:00', '%H:%M')


        # Adjust the time
        if in_time >= datetime.strptime('10:00', '%H:%M') and in_time <= datetime.strptime('11:30', '%H:%M'):
            in_time = datetime.strptime('10:00', '%H:%M')

        if out_time >= datetime.strptime('17:00','%H:%M') and out_time <= datetime.strptime('18:00','%H:%M'):
            out_time = datetime.strptime('18:00','%H:%M')

        # Create Date object
        date = None
        for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%d-%m', '%d/%m', '%d-%m-%y', '%d/%m/%y'):
            try:
                date = datetime.strptime(attendance['Date'], fmt)
                if 'Y' not in fmt.upper():
                    current_year = int(time.strftime('%Y'))
                    date = date.replace(year = current_year)
            except ValueError:
                pass

        if not date:
            print('No valid date format found: ', attendance['Date'])
            date = datetime.strptime('01-01-1990', '%d-%m-%Y')
        
        # Replace the previous date and times
        attendance['In Time'] = in_time.strftime('%H:%M')
        attendance['Out Time'] = out_time.strftime('%H:%M')
        attendance['Date'] = date.strftime('%d-%m-%Y')
        



def export_as_json(attendances):
    with open('attendances.txt', 'w') as f:
        for attendance in attendances:
            json.dump(attendance, f)
            f.write('\n')


def main():
    mail = imaplib.IMAP4_SSL(HOST)
    mail.login(EMAIL, PASSWORD)
    email_messages, sender_names = read_mails(mail)
    attendances = get_details(email_messages, sender_names)
    rectify_attendances(attendances)
    print(attendances)
    export_as_json(attendances)



if __name__ == '__main__':
    main()









