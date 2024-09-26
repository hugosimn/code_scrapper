import imaplib
import email
import csv
import re

def connect_to_email(imap_server, email_address, password):
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_address, password)
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Failed to login: {e}")
        return None

def search_emails(mail, subject):
    mail.select('inbox')
    _, message_numbers = mail.search(None, f'SUBJECT "{subject}"')
    return message_numbers[0].split()

def extract_info(mail, message_number, search_pattern):
    _, msg_data = mail.fetch(message_number, '(RFC822)')
    email_body = msg_data[0][1]
    email_message = email.message_from_bytes(email_body)
    
    # Extract the recipient's email from the "To" field
    recipient = email_message['To']
    recipient_email = re.search(r'<(.+?)>', recipient)
    if recipient_email:
        recipient_email = recipient_email.group(1)
    else:
        recipient_email = recipient.split()[-1]
    
    if email_message.is_multipart():
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode()
                match = re.search(search_pattern, body, re.DOTALL)
                if match:
                    return recipient_email, match.group(1)
    else:
        body = email_message.get_payload(decode=True).decode()
        match = re.search(search_pattern, body, re.DOTALL)
        if match:
            return recipient_email, match.group(1)
    
    return recipient_email, None

def save_to_csv(data, filename):
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerow(['Email', 'Access Code'])
        csv_writer.writerows(data)

def main():
    # Email account settings
    imap_server = 'imap.gmail.com'  # for Gmail accounts
    email_address = 'gmail adress'  # Your email
    password = 'gmail password'  # Your password
    
    # Set search criteria for Coldplay presale
    subject = "Your Coldplay UK 2025 Presale Code"
    search_pattern = r'Your Unique Access Code:\s*(\w+)'
    
    # Connect to email
    mail = connect_to_email(imap_server, email_address, password)
    if mail is None:
        return  # Exit if login failed
    
    # Search for emails
    message_numbers = search_emails(mail, subject)
    print(f"Message numbers found: {message_numbers}")  # Debug line

    # Extract information and store in a list
    results = []
    for num in message_numbers:
        recipient_email, info = extract_info(mail, num, search_pattern)
        print(f"Recipient: {recipient_email}, Info: {info}")  # Debug line
        if info:
            results.append([recipient_email, info])
    
    # Save to CSV if there are results
    if results:
        save_to_csv(results, 'coldplay_presale_codes.csv')
        print(f"Found {len(results)} matching emails. Results saved to coldplay_presale_codes.csv")
    else:
        print("No matching emails found.")

    # Logout
    mail.logout()

if __name__ == "__main__":
    main()
