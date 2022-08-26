#modules
import imaplib
import email
import gspread
import time
ref_num = ""

def verify_amount(email_message):
    for part in email_message.walk():
        if part.get_content_type()=="text/plain" or part.get_content_type()=="text/html":
            message = part.get_payload(decode=True)

            str_msg = message.decode()
            if("$10.00 (CAD)" in str_msg):
                get_reference_num(str_msg)
                return True
            else:
                return False

def get_reference_num(str_msg):
    global ref_num
    ref_num = ""
    if("Reference Number: " in str_msg):
        ref_num = str_msg.split("Reference Number: ")[1]
        ref_num = ref_num.split("\r")[0]

def get_name(email_message):
    name = email_message["from"].split(" <notify@payments.interac.ca>")[0].upper()
    name = name.split(" ")
    if(len(name) == 3):
        name = [name[0], name[2]]
    return name

def get_ref_list():

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
    ref_nums = []
    for num in selected_mails[0].split():
        _, data = mail.fetch(num , '(RFC822)')
        _, bytes_data = data[0]

        #convert the byte data to message
        email_message = email.message_from_bytes(bytes_data)
        if("deposited" in email_message["subject"]):
            if (verify_amount(email_message)):
                ref_nums.append([get_name(email_message),ref_num])
    return ref_nums

def check_if_contains_in_worksheet(wks, ref_nums):
    row_length = wks.row_count
    for i in range(2, row_length):
        if(i == 29):
            time.sleep(70)
        wks_method = wks.acell(f'H{i}').value
        if(wks_method == "E-transfer to test@gmail.com"):
            if(wks.acell(f'J{i}').value == None):
                        wks.update(f'J{i}', 'N')
            wks_ref_num = wks.acell(f'I{i}').value
            wks_names = wks.get(f'K{i}:L{i}')
            for record in ref_nums:
                if(record[1] == wks_ref_num and record[0] == wks_names[0]):
                    wks.update(f'J{i}', 'Y')               
        elif(wks_method == "Cash in person at Jumaah"):
            if(wks.acell(f'J{i}').value == None):
               wks.update(f'J{i}', 'N') 


def main():
    #Get all reference numbers corresponding to the amount
    ref_nums = get_ref_list()
    sa = gspread.service_account()
    sh = sa.open("MSA 2022 Membership Form (Responses)")

    wks = sh.worksheet("Form Responses 1")
    check_if_contains_in_worksheet(wks,ref_nums)

if __name__ == "__main__":
    main()