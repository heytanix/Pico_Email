# Importing necessary libraries for sending emails, handling file formats, and SSL encryption
import smtplib  # To send emails using SMTP
from email.message import EmailMessage  # To construct email messages
import ssl  # For creating a secure SSL context
import csv  # To handle CSV files
import pandas as pd  # To handle Excel (.xlsx) files
import json  # To handle JSON files

# Function to print the file format information to guide the user
def print_file_format_info():
    """
    Prints the required file format for CSV, Excel, or JSON files
    that contain the email, name, and surname details.
    """
    print("\nPlease ensure your file follows this format:")
    
    # Instructions for CSV or Excel files
    print("1. For CSV or Excel (.xlsx):")
    print('   - Column headers must be: "EmailID", "Name", "Surname"')
    print("   Example CSV content:")
    print("   EmailID,Name,Surname")
    print("   john.doe@example.com,John,Doe")
    print("   jane.smith@example.com,Jane,Smith")
    
    # Instructions for JSON files
    print("\n2. For JSON file:")
    print('   - JSON should be an array of dictionaries with keys: "EmailID", "Name", "Surname"')
    print("   Example JSON content:")
    print('   [')
    print('      {"EmailID": "john.doe@example.com", "Name": "John", "Surname": "Doe"},')
    print('      {"EmailID": "jane.smith@example.com", "Name": "Jane", "Surname": "Smith"}')
    print('   ]')

# Getting the sender's email credentials from the user
SENDER_MAIL = input("Enter Yahoo sender email: ")  # Yahoo email address
SENDER_PASSWORD = input("Enter Yahoo sender password: ")  # Yahoo email password

# Getting the email's subject and message template from the user
email_subject = input("Enter the email subject: ")  # Subject of the email
email_body_template = input(
    "Enter the body of the email. Use {name} and {surname} as placeholders: \n"
)  # Email body template where {name} and {surname} will be replaced

# Function to send a single email using the provided recipient details
def send_email(to_email, name, surname):
    """
    Sends an email to a single recipient.
    
    :param to_email: Recipient's email address
    :param name: Recipient's first name
    :param surname: Recipient's surname
    """
    # Create an EmailMessage object to hold the email content
    msg = EmailMessage()
    
    # Fill in the email details: sender, recipient, subject
    msg['From'] = SENDER_MAIL  # Sender email address
    msg['To'] = to_email  # Recipient email address
    msg['Subject'] = email_subject  # Subject of the email
    
    # Format the email body by replacing placeholders with actual names
    message = email_body_template.format(name=name, surname=surname)
    msg.set_content(message)  # Set the email body content
    
    # Create a secure SSL context to encrypt the connection to the Yahoo SMTP server
    context = ssl.create_default_context()
    
    # Sending the email using the Yahoo SMTP server
    with smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465, context=context) as smtp:
        smtp.login(SENDER_MAIL, SENDER_PASSWORD)  # Log in to the Yahoo account
        smtp.send_message(msg)  # Send the email message
        print(f"Email sent to: {name} {surname}, ({to_email})")  # Confirmation message

# Function to handle sending bulk emails based on the selected file format
def send_bulk_emails(file_choice, file_path):
    """
    Sends bulk emails by reading recipient details from a specified file format (CSV, Excel, or JSON).
    
    :param file_choice: The user's choice of file format (1 for CSV, 2 for Excel, 3 for JSON)
    :param file_path: Path to the file containing recipient details
    """
    # Handle CSV file input
    if file_choice == 1:
        with open(file_path, newline='') as file:
            reader = csv.DictReader(file)  # Read the CSV file as a dictionary
            for row in reader:
                send_email(row['EmailID'], row['Name'], row['Surname'])  # Send email to each recipient
    
    # Handle Excel (.xlsx) file input
    elif file_choice == 2:
        data = pd.read_excel(file_path)  # Read the Excel file using pandas
        for index, row in data.iterrows():
            send_email(row['EmailID'], row['Name'], row['Surname'])  # Send email to each recipient
    
    # Handle JSON file input
    elif file_choice == 3:
        with open(file_path, 'r') as file:
            data = json.load(file)  # Load the JSON data
            for entry in data:
                send_email(entry['EmailID'], entry['Name'], entry['Surname'])  # Send email to each recipient
    
    # Invalid file choice
    else:
        print("Invalid file choice. Please enter 1 for CSV, 2 for Excel, or 3 for JSON.")

# Main function to prompt the user for the file type and path
def main():
    """
    Main script function to prompt the user for file type, file path,
    and handle bulk email sending based on the input file.
    """
    # Prompt the user to specify the file type containing the email details
    print("Please specify the type of file containing the details (EmailID, Name, Surname):")
    print("1. CSV (.csv)\n2. Excel (.xlsx)\n3. JSON (.json)\n4. Help")
    
    # Get the user's choice for file format
    file_choice = int(input("Please enter your choice number here: "))
    
    # If the user asks for help, show the file format info
    if file_choice == 4:
        print_file_format_info()  # Show file format instructions
        return

    # If the file choice is invalid, show an error message
    if file_choice not in [1, 2, 3]:
        print("Invalid choice.")
        return
    
    # Prompt the user to enter the file path
    file_path = input("Please enter the file path: ")
    
    # Call the function to send bulk emails based on the selected file format
    send_bulk_emails(file_choice, file_path)

# Run the main script
if __name__ == "__main__":
    main()  # Execute the main function
