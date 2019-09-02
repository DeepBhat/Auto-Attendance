import email

class emailMessage():
    """description of class"""
    def __init__(self, email_message):
        self.email_from = str(email.header.make_header(email.header.decode_header(email_message['From'])))
        angled1 = self.email_from.find('<')
        self.name = self.email_from[0 : angled1 - 1]
        angled2 = self.email_from.find('>')
        self.emailID = self.email_from[angled1 + 1 : angled2]
        self.email_to = str(email.header.make_header(email.header.decode_header(email_message['To'])))
        self.subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))
        self.body = self.get_body(email_message)
        self.body = self.body.replace('\r','')

    def get_body(self, email_message):
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode('utf=8')
                return body