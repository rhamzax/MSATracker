#modules
import imaplib
import email

#credentials
username ="calgarymsa@gmail.com"

#generated app password
app_password = "aiqyflzydlpznhvn"

# https://www.systoolsgroup.com/imap/
gmail_host= 'imap.gmail.com'

#set connection
mail = imaplib.IMAP4_SSL(gmail_host)

#login
mail.login(username, app_password)

#select inbox
mail.select("INBOX")

#select specific mails

_, selected_mails = mail.search(None, '(FROM "(notify@payments.interac.ca")')

#total number of mails from specific user
print("Total Messages from notify@payments.interac.ca:" , len(selected_mails[0].split()))

def verify_amount():
    for part in email_message.walk():
        if part.get_content_type()=="text/plain" or part.get_content_type()=="text/html":
            message = part.get_payload(decode=True)

            str_msg = message.decode()
            if("$855.91 (CAD)" in str_msg):
                return True
            else:
                return False

for num in selected_mails[0].split():
    _, data = mail.fetch(num , '(RFC822)')
    _, bytes_data = data[0]

    #convert the byte data to message
    email_message = email.message_from_bytes(bytes_data)
    print("\n===========================================")
    if (verify_amount()):
        #access data
        print("Subject: ",email_message["subject"])
        print("To:", email_message["to"])
        print("From: ",email_message["from"])
        print("Date: ",email_message["date"])
        for part in email_message.walk():
            if part.get_content_type()=="text/plain" or part.get_content_type()=="text/html":
                message = part.get_payload(decode=True)

                print("Message: \n", message.decode())
                print("==========================================\n")
                break


