#Importing the necessary libraries
import smtplib
from email.message import EmailMessage
import ssl
import csv
import pandas as pd
import json

#Function to print the file format information
def print_file_format_info():
    print("\nPlease ensure your file follows this format:")
    print("1. For CSV or Excel (.xlsx):")
    print('   - Column headers must be: "EmailID", "Name", "Surname"')
    print("   Example CSV content:")
    print("   EmailID,Name,Surname")
    print("   john.doe@example.com,John,Doe")
    print("   jane.smith@example.com,Jane,Smith")
    print("\n2. For JSON file:")
    print('   - JSON should be an array of dictionaries with keys: "EmailID", "Name", "Surname"')
    print("   Example JSON content:")
    print('   [')
    print('      {"EmailID": "john.doe@example.com", "Name": "John", "Surname": "Doe"},')
    print('      {"EmailID": "jane.smith@example.com", "Name": "Jane", "Surname": "Smith"}')
    print('   ]')

#Sender email's credentials (Yahoo)
SENDER_MAIL=input("Enter Yahoo sender email: ")
SENDER_PASSWORD=input("Enter Yahoo sender password: ")

#Getting the email's subject and message from the user
email_subject=input("Enter the email subject: ")
email_body_template=input(
    "Enter the body of the email. Use {name} and {surname} as placeholders: \n"
)

#Defining function to send a single email
def send_email(to_email, name, surname):
    #Creating an email message object
    msg=EmailMessage()
    
    #Filling in the email details
    msg['From']=SENDER_MAIL
    msg['To']=to_email
    msg['Subject']=email_subject
    
    #Creating the email content from user's input
    message=email_body_template.format(name=name, surname=surname)
    msg.set_content(message)
    
    #Setting up the server using SSL encryption (Yahoo)
    context=ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465, context=context) as smtp:
        smtp.login(SENDER_MAIL, SENDER_PASSWORD)
        smtp.send_message(msg)
        print(f"Email sent to: {name} {surname}, ({to_email})")

#Function to handle sending emails based on file type
def send_bulk_emails(file_choice, file_path):
    if file_choice==1:
        with open(file_path, newline='') as file:
            reader=csv.DictReader(file)
            for row in reader:
                send_email(row['EmailID'], row['Name'], row['Surname'])
    elif file_choice==2:
        data=pd.read_excel(file_path)
        for index, row in data.iterrows():
            send_email(row['EmailID'], row['Name'], row['Surname'])
    elif file_choice==3:
        with open(file_path, 'r') as file:
            data=json.load(file)
            for entry in data:
                send_email(entry['EmailID'], entry['Name'], entry['Surname'])
    else:
        print("Invalid file choice. Please enter 1 for CSV, 2 for Excel, or 3 for JSON.")

#Main script to get file type from the user
def main():
    print("Please specify the type of file containing the details (EmailID, Name, Surname):")
    print("1. CSV (.csv)\n2. Excel (.xlsx)\n3. JSON (.json)\n4. Help")
    
    file_choice=int(input("Please enter your choice number here: "))
    
    if file_choice==4:
        print_file_format_info()
        return

    if file_choice not in [1, 2, 3]:
        print("Invalid choice.")
        return
    
    file_path=input("Please enter the file path: ")
    
    #Function call - to send bulk emails
    send_bulk_emails(file_choice, file_path)

#Run the script
main()