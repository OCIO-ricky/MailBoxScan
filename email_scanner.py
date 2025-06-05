import os
import csv
from dotenv import load_dotenv
from imap_tools import MailBox, AND
import re
from email.utils import parseaddr
from datetime import datetime
import traceback
from msal import ConfidentialClientApplication # For Service Principal OAuth

# Load environment variables from .env file
load_dotenv()

# Configuration variables
IMAP_SERVER = os.getenv("IMAP_SERVER")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
# EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD") # Not used for OAuth
TARGET_SUBJECT = os.getenv("TARGET_SUBJECT")
SEARCH_QUESTION = os.getenv("SEARCH_QUESTION")
OUTPUT_CSV_FILE = os.getenv("OUTPUT_CSV_FILE")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

def extract_answer(text_body, question):
    """
    Extracts 'Yes' or 'No' answer following the specified question in the text.
    Searches within a limited window after the question.
    """
    if not text_body or not question:
        return None

    text_body_lower = text_body.lower()
    question_lower = question.lower()

    question_index = text_body_lower.find(question_lower)
    if question_index == -1:
        return None  # Question not found

    # Define a window (e.g., next 100 characters) after the question to search for an answer
    search_window_start = question_index + len(question)
    # Ensure the window does not exceed text length
    search_window_end = min(search_window_start + 100, len(text_body_lower))
    text_after_question = text_body_lower[search_window_start:search_window_end]

    # Search for "yes" or "no" (whole word) in the window.
    # text_after_question is already lowercase.
    match_yes = re.search(r'\byes\b', text_after_question)
    match_no = re.search(r'\bno\b', text_after_question)

    if match_yes and match_no:
        # If both "yes" and "no" are found, prioritize the one that appears first
        if match_yes.start() < match_no.start():
            return "Yes"
        else:
            return "No"
    elif match_yes:
        return "Yes"
    elif match_no:
        return "No"

    return None  # No clear Yes/No found in the window

def scan_emails():
    """
    Connects to the Outlook mailbox, scans emails, and extracts information.
    """
    results = []
    print("Attempting to log in to Outlook...")
    try:
        # --- OAuth 2.0 Token Acquisition for Service Principal ---
        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        app = ConfidentialClientApplication(
            CLIENT_ID,
            authority=authority,
            client_credential=CLIENT_SECRET
        )

        # Scope for Exchange Online IMAP using client credentials
        scopes = ["https://outlook.office365.com/.default"]

        print("Acquiring OAuth token for service principal...")
        token_result = app.acquire_token_for_client(scopes=scopes)

        if "access_token" not in token_result:
            error_message = "Failed to acquire OAuth token.\n"
            error_message += f"Error: {token_result.get('error')}\n"
            error_message += f"Error description: {token_result.get('error_description')}\n"
            error_message += f"Correlation ID: {token_result.get('correlation_id')}\n"
            error_message += "Please check your .env file for TENANT_ID, CLIENT_ID, CLIENT_SECRET, and Azure AD app registration permissions (IMAP.AccessAsApp and admin consent) and mailbox permissions for the service principal."
            print(error_message)
            return results # Or raise an exception

        access_token = token_result["access_token"]
        print("OAuth token acquired successfully.")
        # --- End OAuth 2.0 Token Acquisition ---

        with MailBox(IMAP_SERVER).login_oauth(EMAIL_ADDRESS, access_token) as mailbox:
            print(f"Successfully logged in to {EMAIL_ADDRESS} on {IMAP_SERVER} using OAuth.")
            print(f"Current folder: {mailbox.folder.get()}") # Usually INBOX by default
            
            # If your emails are in a specific folder other than INBOX, uncomment and set it:
            # mailbox.folder.set('YourSpecificFolderName')
            # print(f"Changed folder to: {mailbox.folder.get()}")

            print(f"Searching for emails with subject containing: '{TARGET_SUBJECT}'...")

            emails_inspected_count = 0
            emails_with_answer_count = 0
            any_emails_fetched_with_subject = False

            # Fetch emails matching the subject.
            # `bulk=True` can improve performance for many messages.
            # `charset='UTF-8'` helps with character encoding.
            # `imap_tools` `msg.text` prefers plain text, falls back to html-to-text conversion.
            for msg in mailbox.fetch(AND(subject=TARGET_SUBJECT), charset='UTF-8', bulk=True):
                any_emails_fetched_with_subject = True
                emails_inspected_count += 1

                print(f"\nProcessing email {emails_inspected_count}:")
                print(f"  UID: {msg.uid}, Subject: '{msg.subject}'")
                print(f"  From: {msg.from_}, Date: {msg.date}")

                sender_name, sender_email = parseaddr(msg.from_)
                received_date_str = msg.date.strftime("%Y-%m-%d %H:%M:%S") if isinstance(msg.date, datetime) else str(msg.date)

                email_body = msg.text  # msg.text provides plain text or HTML-converted-to-text

                if not email_body:
                    print(f"  - Email UID {msg.uid} has no textual body content. Skipping.")
                    continue

                answer = extract_answer(email_body, SEARCH_QUESTION)

                if answer:
                    print(f"  - Found: Sender Name='{sender_name}', Sender Email='{sender_email}', Date='{received_date_str}', Answer='{answer}'")
                    results.append({
                        "Sender Name": sender_name or "",
                        "Sender Email": sender_email or "",
                        "Date Received": received_date_str,
                        "Answer": answer
                    })
                    emails_with_answer_count += 1
                else:
                    print(f"  - Question '{SEARCH_QUESTION}' or a Yes/No answer not found in email UID {msg.uid}.")
            
            if not any_emails_fetched_with_subject:
                print(f"\nNo emails found with subject containing: '{TARGET_SUBJECT}'.")
            else:
                print(f"\nFinished processing. Inspected {emails_inspected_count} email(s) with matching subject.")
                print(f"Found {emails_with_answer_count} email(s) containing the question and a Yes/No answer.")

    except Exception as e:
        print(f"\nAn error occurred during email processing: {e}")
        print("Detailed error information:")
        traceback.print_exc()
    
    return results

def save_to_csv(data, filename):
    """
    Saves the extracted data to a CSV file.
    """
    if not data:
        print("No data to save to CSV.")
        return

    print(f"\nSaving data to {filename}...")
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ["Sender Name", "Sender Email", "Date Received", "Answer"]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            for row in data:
                writer.writerow(row)
        print(f"Data successfully saved to {filename}")
    except IOError as e:
        print(f"Error writing to CSV file {filename}: {e}")
    except Exception as e:
        print(f"An unexpected error occurred while saving CSV: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    if not all([IMAP_SERVER, EMAIL_ADDRESS, TENANT_ID, CLIENT_ID, CLIENT_SECRET, TARGET_SUBJECT, SEARCH_QUESTION, OUTPUT_CSV_FILE]):
        print("Error: Critical configuration variables are missing in the .env file.")
        print("Please ensure IMAP_SERVER, EMAIL_ADDRESS, TENANT_ID, CLIENT_ID, CLIENT_SECRET, TARGET_SUBJECT, SEARCH_QUESTION, and OUTPUT_CSV_FILE are set for OAuth authentication.")
    else:
        print("Starting email scan script...")
        extracted_data = scan_emails()
        if extracted_data:
            save_to_csv(extracted_data, OUTPUT_CSV_FILE)
        else:
            print("No data was extracted that matched the criteria.")
        print("\nScript finished.")