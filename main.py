#modules
import imaplib
import email
import gspread

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
    search_range = f'I2:I{row_length}'
    wks_ref_nums = wks.get(search_range)

    search_range = f'B2:C{row_length}'
    wks_names = wks.get(search_range)
    
    for wks_num in wks_ref_nums:
        for record in ref_nums:
            if(record[1] == wks_num[0]):
                print(record)
        
def update_worksheet():
    pass

def main():
    #Get all reference numbers corresponding to the amount
    ref_nums = [[['KINDA', 'CHAKAS'], 'CAGncAVW'], [['NOUR', 'ALMRIRI'], 'CAj6fMcr'], [['TAHSEEN', 'MONEM'], 'CAz28x2Y'], [['MALAIKA', 'KHAN'], 'CAaWwcQN'], [['HAMZA', 'AFZAAL'], 'CA7aMC9u'], [['SHAMIS', 'ALI'], 'CA2hhhsy'], [['FATIMA', 'WARRAICH'], 'CAdJkqYP'], [['REZWAN', 'AHMED'], 'CA3fZWBU']]
    sa = gspread.service_account()
    sh = sa.open("MSA 2022 Membership Form (Responses)")

    wks = sh.worksheet("Form Responses 1")
    check_if_contains_in_worksheet(wks,ref_nums)

if __name__ == "__main__":
    main()