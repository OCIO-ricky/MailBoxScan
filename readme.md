# Outlook Email Survey Scanner

This Python script automates the process of scanning a specified Microsoft 365 Outlook mailbox for emails matching a particular subject. It extracts "Yes" or "No" answers to a defined question within the email body, records these responses in a CSV file, and moves processed emails to a designated subfolder in the mailbox.

## Features

-   **Targeted Email Scanning**: Searches for emails based on a specific subject line.
-   **Answer Extraction**: Parses email bodies to find a specific question and extracts a "Yes" or "No" answer immediately following it.
-   **CSV Reporting**:
    -   Saves extracted sender information, date received, the answer, and a "Last Updated" timestamp to a CSV file.
    -   Stores CSV files in an `output/` directory.
    -   Updates existing records in the CSV if a newer email from the same sender provides a different answer.
    -   Backs up the previous CSV file with a timestamp before writing changes.
-   **Email Management**: Moves emails processed (either new or updated) to a specified subfolder in Outlook.
-   **Comprehensive Logging**:
    -   Logs script activity, informational messages, warnings, and errors to both the console and a log file (`logs/email_scanner.log`).
 -   **Configuration via `.env` file**: Manages sensitive credentials and settings securely.

## Prerequisites

**Microsoft Graph API Integration**: Uses the `O365` library to interact with Microsoft Graph API for modern and secure access to mailbox data.
**OAuth 2.0 Authentication**: Employs client credentials flow (application authentication) for secure, unattended access.

1.  **Python**: Python 3.7+ is recommended.
2.  **Microsoft 365 Account**: A mailbox to scan.
3.  **Azure AD App Registration**:
    -   An application registered in Azure Active Directory.
    -   **Required API Permissions** (Microsoft Graph, Application type):
        -   `Mail.Read`: To read emails from the mailbox.
        -   `Mail.ReadWrite`: To move emails to a subfolder.
    -   Admin consent granted for these permissions.
    -   A Client ID and Client Secret generated for the application.
    -   Your Azure AD Tenant ID.
4.  **Python Libraries**:
    -   `O365`
    -   `python-dotenv`
    (These can be installed via `pip`)

## Setup

1.  **Clone the Repository (if applicable)**:
    ```bash
    git clone <repository-url>
    cd MailBoxScan
    ```

2.  **Install Dependencies**:
    It's recommended to create a `requirements.txt` file:
    ```text
    # requirements.txt
    O365
    python-dotenv
    ```
    Then install using pip:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Configure Environment Variables**:
    -   Copy the `.env.template` file to a new file named `.env` in the root of the project:
        ```bash
        cp .env.template .env
        ```
    -   Edit the `.env` file and fill in your specific details for all the variables.

## Configuration (`.env` file)

The `.env` file contains the following configuration variables:

-   `EMAIL_ADDRESS`: The email address of the target mailbox to scan (e.g., `survey@example.com`).
-   `TENANT_ID`: Your Azure AD Tenant ID.
-   `CLIENT_ID`: The Application (Client) ID of your Azure AD registered app.
-   `CLIENT_SECRET`: The Client Secret value for your Azure AD registered app.
-   `TARGET_SUBJECT`: The subject line (or a part of it) the script will look for in emails.
-   `SEARCH_QUESTION`: The exact question text the script will search for within the email body to find a "Yes/No" answer.
-   `OUTPUT_CSV_FILE`: The filename for the CSV report (e.g., `survey_results.csv`). This file will be stored in the `output/` directory.
-   `PROCESSED_FOLDER_NAME`: The name of the subfolder in the Outlook mailbox where emails with processed answers will be moved (e.g., `ProcessedSurveyResponses`).

## Running the Script

Once the setup and configuration are complete, you can run the script from the project's root directory:

```bash
python email_scanner.py
```

## Output

1.  **CSV File**:
    -   Located in the `output/` directory (e.g., `output/mobile_phone_survey_results.csv`).
    -   Contains columns: `Sender Name`, `Sender Email`, `Date Received`, `Answer`, `Last Updated`.
    -   If changes are made to an existing CSV, the old version is backed up in the same `output/` directory with a timestamp appended to its name.

2.  **Log File**:
    -   Located at `logs/email_scanner.log`.
    -   Contains detailed logs of the script's execution, including informational messages, warnings, and errors.
    -   New log entries are appended to this file on each run.
    -   Console output also provides real-time logging.

3.  **Processed Emails in Outlook**:
    -   Emails from which an answer was successfully extracted and recorded (or updated) are moved to the subfolder specified by `PROCESSED_FOLDER_NAME` in the target Outlook mailbox.
