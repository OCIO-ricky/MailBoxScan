import os
import csv
from dotenv import load_dotenv
import logging
import re
from datetime import datetime, timezone
import traceback
from O365 import Account, MSGraphProtocol

# Load environment variables from .env file
load_dotenv()

# --- Logging Setup ---
LOG_DIR = "logs"
LOG_FILE = os.path.join(LOG_DIR, "email_scanner.log")

# Ensure log directory exists
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# Configure root logger
# In a production environment, you might want a more sophisticated logging setup
# (e.g., logging to a file, rotating logs, different formats per handler).
logger = logging.getLogger() # Get the root logger
logger.setLevel(logging.INFO) # Set the default level for the root logger

# Console Handler (keeps existing console output behavior)
console_handler = logging.StreamHandler()
console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

# File Handler
file_handler = logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8') # Append mode
file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(lineno)d - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[console_handler, file_handler] # Use configured handlers
)

# Configuration variables
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
TARGET_SUBJECT = os.getenv("TARGET_SUBJECT")
SEARCH_QUESTION = os.getenv("SEARCH_QUESTION")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
PROCESSED_FOLDER_NAME = os.getenv("PROCESSED_FOLDER_NAME")
CSV_BACKUP_TIMESTAMP_FORMAT = "_%Y%m%d_%H%M%S"

OUTPUT_DIR = "output"
OUTPUT_CSV_FILENAME = os.getenv("OUTPUT_CSV_FILE") # Get the filename from .env
OUTPUT_CSV_FILE_PATH = os.path.join(OUTPUT_DIR, OUTPUT_CSV_FILENAME) if OUTPUT_CSV_FILENAME else None

# Ensure output directory exists
if OUTPUT_DIR and not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)


class CustomMSGraphProtocol(MSGraphProtocol):
    """Custom protocol to set default headers for Graph API requests."""
    def get_session(self, **kwargs):
        session = super().get_session(**kwargs)
        session.headers.update({
            'Prefer': 'outlook.body-type="text"',
            'ConsistencyLevel': 'eventual' # Often required for $filter, $search, $count
        })
        logging.debug("CustomMSGraphProtocol: Session headers updated.")
        return session

DEFAULT_PROTOCOL = CustomMSGraphProtocol()

def extract_answer(text_body, question):
    """
    Extracts 'Yes' or 'No' answer following the specified question in the text.
    Searches within a limited window after the question.
    """
    if not text_body or not question:
        logging.warning("extract_answer called with empty text_body or question.")
        return None

    text_body_lower = text_body.lower()
    question_lower = question.lower()

    question_index = text_body_lower.find(question_lower)
    if question_index == -1:
        logging.debug(f"Question '{question}' not found in text body.")
        return None  # Question not found

    search_window_start = question_index + len(question)
    search_window_end = min(search_window_start + 100, len(text_body_lower))
    text_after_question = text_body_lower[search_window_start:search_window_end]

    match_yes = re.search(r'\byes\b', text_after_question)
    match_no = re.search(r'\bno\b', text_after_question)

    if match_yes and match_no:
        if match_yes.start() < match_no.start():
            logging.debug(f"Found 'Yes' (before 'No') for question: '{question}'")
            return "Yes"
        else:
            logging.debug(f"Found 'No' (before 'Yes') for question: '{question}'")
            return "No"
    elif match_yes:
        logging.debug(f"Found 'Yes' for question: '{question}'")
        return "Yes"
    elif match_no:
        logging.debug(f"Found 'No' for question: '{question}'")
        return "No"

    logging.debug(f"No clear Yes/No answer found for question: '{question}'")
    return None

def scan_emails():
    """
    Connects to Outlook via Microsoft Graph API, scans emails, and extracts information.
    Handles login/token acquisition before attempting to scan.
    """
    # Use a dictionary to store/update records by sender_email for efficient updates
    # Key: sender_email (lowercase), Value: {"Sender Name": ..., "Sender Email": ..., "Date Received": ..., "Answer": ..., "Last Updated": datetime_object}
    current_records = {}
    data_changed_during_scan = False # Flag to track if any record was added or updated
    logging.info("Attempting to authenticate with Microsoft Graph API via O365 library...")

    # --- Authentication and Account Setup using O365 library ---
    try:
        credentials = (CLIENT_ID, CLIENT_SECRET)
        account = Account(credentials, auth_flow_type='credentials', tenant_id=TENANT_ID, protocol=DEFAULT_PROTOCOL)

        if not account.is_authenticated:
            logging.info("Account not yet authenticated. Attempting authentication...")
            if not account.authenticate(scopes=['https://graph.microsoft.com/.default']):
                logging.critical(
                    "Failed to authenticate with O365 library. "
                    "Please check .env for TENANT_ID, CLIENT_ID, CLIENT_SECRET, "
                    "Azure AD app permissions (e.g., Mail.Read - Application), and admin consent. "
                    "Also, verify application access policies if applicable."
                )
                return current_records, data_changed_during_scan # Return empty records and no changes
        
        logging.info("Authenticated successfully. Ready to scan emails.")

    except Exception as e:
        logging.critical(f"An unexpected error occurred during Microsoft Graph API authentication setup: {e}. Exiting.", exc_info=False)
        return current_records, data_changed_during_scan # Return empty records and no changes

    # --- Load existing CSV data ---
    logging.info(f"Attempting to load previous responses from '{OUTPUT_CSV_FILE_PATH}' if it exists...")
    if OUTPUT_CSV_FILE_PATH and os.path.exists(OUTPUT_CSV_FILE_PATH):
        try:
            with open(OUTPUT_CSV_FILE_PATH, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    sender_email_lc = row.get("Sender Email", "").lower()
                    if sender_email_lc:
                        # Store the last updated timestamp as a datetime object for comparison
                        try:
                            # Assuming "Last Updated" is stored in ISO format or similar parsable format
                            last_updated_dt = datetime.fromisoformat(row.get("Last Updated").replace("Z", "+00:00"))
                        except (ValueError, AttributeError):
                            # Fallback if "Last Updated" is missing or not parsable, treat as old
                            last_updated_dt = datetime.min.replace(tzinfo=timezone.utc)
                        row_data = row.copy()
                        row_data["Last Updated DT"] = last_updated_dt # Store actual datetime object
                        current_records[sender_email_lc] = row_data
            logging.info(f"Loaded {len(current_records)} existing records from {OUTPUT_CSV_FILE_PATH}.")
        except Exception as e:
            logging.error(f"Could not read existing CSV {OUTPUT_CSV_FILE_PATH}. Will proceed as if creating a new one. Error: {e}", exc_info=False)
            current_records = {} # Start fresh if read fails

    # --- Email Scanning using O365 library ---
    try:
        logging.info(f"Searching for emails in '{EMAIL_ADDRESS}' with subject containing: '{TARGET_SUBJECT}'")
        
        # Get or create the folder for processed emails
        # Note: This requires Mail.ReadWrite permissions for the application.
        processed_folder = None
        if PROCESSED_FOLDER_NAME:
            logging.info(f"Attempting to get or create processed folder: '{PROCESSED_FOLDER_NAME}'")
            mailbox_for_folders = account.mailbox(resource=EMAIL_ADDRESS) # Ensure we use the correct mailbox context
            processed_folder = mailbox_for_folders.get_folder_by_name(PROCESSED_FOLDER_NAME)
            if not processed_folder:
                logging.info(f"Folder '{PROCESSED_FOLDER_NAME}' not found. Attempting to create it.")
                # new_folder creates it at the root of the mailbox (same level as Inbox)
                processed_folder = mailbox_for_folders.new_folder(PROCESSED_FOLDER_NAME)
                logging.info(f"Folder '{PROCESSED_FOLDER_NAME}' created successfully.")
            else:
                logging.info(f"Successfully retrieved folder '{PROCESSED_FOLDER_NAME}'.")
        else:
            logging.warning("PROCESSED_FOLDER_NAME is not set. Emails will not be moved.")

        mailbox = account.mailbox(resource=EMAIL_ADDRESS)

        odata_filter = f"contains(subject, '{TARGET_SUBJECT}')"
        select_fields = ['id', 'subject', 'from', 'receivedDateTime', 'body']
        
        messages_generator = mailbox.get_messages(
            query=odata_filter, 
            select=select_fields, 
            limit=None, # Fetch all matching messages
            download_attachments=False
        )

        emails_inspected_count = 0
        emails_with_answer_count = 0
        any_messages_found = False
        # processed_sender_answers set is no longer needed with the new update logic

        for msg in messages_generator:
            any_messages_found = False
            emails_inspected_count += 1
            logging.info(f"Processing email {emails_inspected_count}: ID='{msg.object_id}', Subject='{msg.subject}'")

            sender_name = msg.sender.name if msg.sender else "N/A"
            sender_email = msg.sender.address if msg.sender else "N/A"
            
            received_dt = msg.received
            # Store dates in ISO format for easier parsing and consistency
            # O365's received_dt is already timezone-aware (usually UTC)
            received_date_iso = received_dt.isoformat() if received_dt else None
            received_date_str = received_dt.strftime("%Y-%m-%d %H:%M:%S %Z") if received_dt else "N/A" # For display/CSV if preferred
            logging.debug(f"  From: {sender_name} <{sender_email}>, Date: {received_date_str}")

            email_body = msg.body
            if not email_body:
                logging.warning(f"Email ID {msg.object_id} has no textual body content. Skipping.")
                continue

            answer = extract_answer(email_body, SEARCH_QUESTION)
            if answer:
                sender_email_lc = sender_email.lower()
                new_record_data = {
                    "Sender Name": sender_name or "",
                    "Sender Email": sender_email, # Store original case for CSV
                    "Date Received": received_date_str, # Or received_date_iso for consistency
                    "Answer": answer,
                    "Last Updated": received_date_iso, # ISO format string
                    "Last Updated DT": received_dt # datetime object for comparison
                }

                if sender_email_lc in current_records:
                    existing_record = current_records[sender_email_lc]
                    # Compare based on the datetime object of the email
                    if received_dt and received_dt > existing_record.get("Last Updated DT", datetime.min.replace(tzinfo=timezone.utc)):
                        logging.info(f"Updating record for sender '{sender_email}'. Old answer: '{existing_record.get('Answer')}' on {existing_record.get('Last Updated')}. New answer: '{answer}' on {received_date_iso} (Email ID {msg.object_id}).")
                        current_records[sender_email_lc] = new_record_data
                        data_changed_during_scan = True
                        emails_with_answer_count += 1 # Count as an update
                        # Move email if it resulted in an update
                        if processed_folder:
                            try:
                                logging.info(f"Moving updated email ID {msg.object_id} to folder '{PROCESSED_FOLDER_NAME}'.")
                                msg.move(processed_folder)
                            except Exception as move_err:
                                logging.error(f"Failed to move email ID {msg.object_id}: {move_err}", exc_info=False)
                    else:
                        logging.info(f"Existing record for sender '{sender_email}' is more recent or same. New email (ID {msg.object_id}) with answer '{answer}' on {received_date_iso} not processed as update.")
                        # Optionally move this older/same-date email if it also has the target subject, even if not updating CSV
                        # This depends on desired behavior for emails that don't change the CSV state.
                        # For now, we only move if it *updates* the CSV record.
                else:
                    logging.info(f"Adding new record for sender '{sender_email}' with answer '{answer}' on {received_date_iso} (Email ID {msg.object_id}).")
                    current_records[sender_email_lc] = new_record_data
                    data_changed_during_scan = True
                    emails_with_answer_count += 1
                    # Move email if it's a new record
                    if processed_folder:
                        try:
                            logging.info(f"Moving new record email ID {msg.object_id} to folder '{PROCESSED_FOLDER_NAME}'.")
                            msg.move(processed_folder)
                        except Exception as move_err:
                            logging.error(f"Failed to move email ID {msg.object_id}: {move_err}", exc_info=False)
            else:
                logging.debug(f"  Question or Yes/No answer not found in email ID {msg.object_id}.")
        
        if not emails_with_answer_count and not any_messages_found: # If no emails were even found with the subject
            logging.info(f"No emails found with subject containing: '{TARGET_SUBJECT}' in the target mailbox.")
        elif not emails_with_answer_count and any_messages_found: # Emails found, but none had answers or led to updates/new records
             logging.info(f"Processed {emails_inspected_count} email(s) with matching subject, but no new/updated answers were recorded.")
        else:
            logging.info(f"Finished processing. Inspected {emails_inspected_count} email(s) with matching subject.")
            logging.info(f"Found {emails_with_answer_count} email(s) containing the question and a Yes/No answer.")

    except Exception as e:
        logging.critical(f"An error occurred during email scanning: {e}", exc_info=False)
     
    return current_records, data_changed_during_scan

def save_to_csv(records_to_save, filename, data_was_changed):
    """
    Saves the records to a CSV file.
    If data_was_changed is True and the file exists, it backs up the old file.
    """
    if not records_to_save:
        logging.info("No records to save to CSV.")
        return

    if not filename:
        logging.error("Output CSV filepath is not configured. Cannot save CSV.")
        return

    if not data_was_changed and os.path.exists(filename): # filename is now OUTPUT_CSV_FILE_PATH
        logging.info(f"No changes detected in data. CSV file '{filename}' will not be modified.")
        return

    # Backup existing file if it exists and data has changed
    if os.path.exists(filename) and data_was_changed:
        try:
            backup_filename = f"{os.path.splitext(filename)[0]}{datetime.now().strftime(CSV_BACKUP_TIMESTAMP_FORMAT)}{os.path.splitext(filename)[1]}"
            os.rename(filename, backup_filename)
            logging.info(f"Backed up existing CSV file to '{backup_filename}'.")
        except Exception as backup_err:
            logging.error(f"Could not back up existing CSV file '{filename}'. Error: {backup_err}", exc_info=False)
            # Decide if you want to proceed with overwrite or stop. For now, we'll proceed.

    logging.info(f"Saving {len(records_to_save)} records to {filename}...")
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            # "Date Received" now reflects the date of the email that provided the latest answer.
            # "Last Updated" is the timestamp of that latest email in ISO format.
            fieldnames = ["Sender Name", "Sender Email", "Date Received", "Answer", "Last Updated"]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            # Sort by sender email for consistent output, optional
            for record_key in sorted(records_to_save.keys()):
                record = records_to_save[record_key]
                # Prepare row for CSV, ensuring only defined fieldnames are written
                row_to_write = {field: record.get(field, "") for field in fieldnames}
                writer.writerow(row_to_write)
        logging.info(f"Data successfully saved to {filename}")
    except IOError as e:
        logging.critical(f"Error writing to CSV file {filename}: {e}", exc_info=False)
    except Exception as e:
        logging.critical(f"An unexpected error occurred while saving CSV: {e}", exc_info=False)
def main():
    """
    Main function to orchestrate the email scanning and data saving process.
    """
    required_env_vars_map = {
        "EMAIL_ADDRESS": EMAIL_ADDRESS, "TENANT_ID": TENANT_ID,
        "CLIENT_ID": CLIENT_ID, "CLIENT_SECRET": CLIENT_SECRET,
        "TARGET_SUBJECT": TARGET_SUBJECT, "SEARCH_QUESTION": SEARCH_QUESTION, 
        "PROCESSED_FOLDER_NAME": PROCESSED_FOLDER_NAME,
        "OUTPUT_CSV_FILE": OUTPUT_CSV_FILENAME # Check for the filename, path is constructed
    }
    
    missing_vars = [name for name, value in required_env_vars_map.items() if not value]
    
    if missing_vars:
        logging.critical(f"Critical configuration variables are missing in the .env file: {', '.join(missing_vars)}")
        logging.critical("Please ensure all required variables are set.")
        logging.critical("For Microsoft Graph API (using O365 library), ensure your Azure AD app registration has the necessary permissions (e.g., Mail.Read and Mail.ReadWrite for Application if moving emails) and admin consent.")
        return

    logging.info("Starting email scan script using O365 library for Microsoft Graph API...")
    final_records, data_changed = scan_emails()
    
    if OUTPUT_CSV_FILE_PATH: # Ensure path is configured before trying to save
        if final_records or data_changed: # Save if there are records or if data changed (e.g. all records removed)
            save_to_csv(final_records, OUTPUT_CSV_FILE_PATH, data_changed)
        else:
            logging.info("No records to process or save after scanning, and no data changes detected.")
    else:
        logging.error("OUTPUT_CSV_FILE is not defined in .env. Cannot save results.")
    
    logging.info("Script finished.")

if __name__ == "__main__":
    main()
