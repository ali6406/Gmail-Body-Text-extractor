import imaplib
import email
import re
from dateutil import parser
from email.header import decode_header
import tkinter as tk
from tkinter import ttk
import os

def decode_subject(subject):
    decoded_subject = []
    for s, encoding in decode_header(subject):
        if isinstance(s, bytes):
            decoded_subject.append(s.decode(encoding or 'utf-8'))
        else:
            decoded_subject.append(s)
    return ' '.join(decoded_subject)

def fetch_and_display_emails():
    sender_email = sender_entry.get()
    target_year = int(year_entry.get())
    target_month = selected_month.get()  # Get the selected month from the StringVar

    with open("output.txt", "w", encoding="utf-8") as output_file:
        imap_url = 'imap.gmail.com'   # Enable IMAP in Gmail settings to make it work
        my_mail = imaplib.IMAP4_SSL(imap_url)
        user = "your.mail@gmail.com"  
        password = "abcd efgh uxyz"  # Set an App password in Gmail settings to get this
        my_mail.login(user, password)
        my_mail.select('Inbox')

        key = 'FROM'
        value = sender_email
        _, data = my_mail.search(None, key, value)

        mail_id_list = data[0].split()
        msgs = []
        for num in mail_id_list:
            typ, data = my_mail.fetch(num, '(RFC822)')
            msgs.append(data)

        for msg in msgs[::-1]:
            for response_part in msg:
                if type(response_part) is tuple:
                    my_msg = email.message_from_bytes((response_part[1]))
                    email_date = parser.parse(my_msg['Date'])
                    email_year = email_date.year
                    email_month = email_date.month
                    if email_year == target_year and email_month == months.index(target_month) + 1:  # Adjust index to match month number
                        output_file.write("_________________________________________\n")
                        output_file.write(f"Date: {my_msg['Date']}\n")
                        output_file.write(f"Subject: {decode_subject(my_msg['subject'])}\n")
                        output_file.write(f"From: {my_msg['from']}\n")
                        output_file.write("\n")
                        for part in my_msg.walk():
                            if part.get_content_type() == 'text/plain':
                                body_text = part.get_payload(decode=True).decode('utf-8')
                                patterns_to_remove = [
                                    r'See all jobs on LinkedIn:.*'
                                ]
                                for pattern in patterns_to_remove:
                                    match = re.search(pattern, body_text)
                                    if match:
                                        body_text = body_text[:match.start()]
                                output_file.write(body_text.strip() + "\n")
                                break

    with open("output.txt", "r", encoding="utf-8") as output_file:
        email_data = output_file.read()
        output_text.delete(1.0, tk.END)
        output_text.insert(tk.END, email_data)

# Store emails in saved_emails.txt
def save_email_address():
    email_address = sender_entry.get()
    filename = "saved_emails.txt"
    # Check if the file exists
    if not os.path.exists(filename):
        # Create the file if it doesn't exist
        with open(filename, "w"):
            pass
    # Save the email address to the file
    with open(filename, "a") as file:
        file.write(email_address + "\n")


def load_saved_emails():
    filename = "saved_emails.txt"
    if os.path.exists(filename):
        with open(filename, "r") as file:
            return file.readlines()
    else:
        return []

# Create the main Tkinter window in full screen
root = tk.Tk()
root.title("Email Fetcher")

# Set window size and position to fullscreen without covering taskbar
root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))

sender_label = tk.Label(root, text="Sender email:")
sender_label.grid(row=0, column=1, pady=10)

# Create a Frame to contain the Combobox widget and the Entry widget
sender_frame = tk.Frame(root)
sender_frame.grid(row=0, column=1, padx=10, pady=10)

# Load saved email addresses
saved_emails = load_saved_emails()

# Sender Email Combobox
sender_combobox = ttk.Combobox(sender_frame, width=40, values=saved_emails)
sender_combobox.pack(side="left")

# Sender Email Entry
sender_entry = tk.Entry(sender_frame, width=40)
sender_entry.pack(side="left")

# "+" Button to save email address
add_email_button = tk.Button(sender_frame, text="+", command=save_email_address)
add_email_button.pack(side="right")

# Target Year
year_label = tk.Label(root, text="Target Year:")
year_label.grid(row=0, column=2, pady=10)

year_entry = tk.Entry(root)
year_entry.grid(row=0, column=3, pady=10)

# Target Month
month_label = tk.Label(root, text="Target Month:")
month_label.grid(row=0, column=4, pady=10)

months = ["January", "February", "March", "April", "May", "June","July", "August", "September", "October", "November", "December"]

# Create a StringVar to store the selected month
selected_month = tk.StringVar(root)
selected_month.set("January")  # Set default value

# Create the OptionMenu widget
month_option = tk.OptionMenu(root, selected_month, *months)
month_option.grid(row=0, column=5, pady=10)

# Fetch Button
fetch_button = tk.Button(root, text="Fetch Emails", command=fetch_and_display_emails)
fetch_button.grid(row=0, column=6, pady=10)

# Output Text Box with Scrollbar
output_text = tk.Text(root, width=175, height=40)
output_text.grid(row=2, column=1, columnspan=6, padx=5, pady=25)

scrollbar = tk.Scrollbar(root, command=output_text.yview)
scrollbar.grid(row=2, column=7, sticky='ns', padx=0, pady=25)
output_text['yscrollcommand'] = scrollbar.set

root.mainloop()
