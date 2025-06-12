We are initiating an effort to identify CDC staff who currently have a government-provided mobile device but may no longer need it. As part of this initiative, an agency-wide email will be sent asking recipients to confirm whether they still require their government-issued mobile phone. Respondents will be asked to reply using a specific subject line and indicate "Yes" or "No" in the body of the email.
To streamline the collection and processing of these responses, we have developed a Python-based script that scans an Outlook mailbox for replies with the designated subject line, captures the sender’s name and email, and records those who respond with “No” in a structured .csv file.

## How to Set Up and Run the Script

Clone the Repository: First, clone the project repository to your local machine:

```bash
git clone https://github.com/OCIO-ricky/MailBoxScan.git
```

```bash
cd MailBoxScan
```

Create the Directory (Alternative): If not cloning, make a new folder for your project (e.g., outlook_scanner).

Save the Files: If you haven't cloned the repository, ensure `email_scanner.py`, `requirements.txt`, and `.env.template` are in your project directory.

Configure Environment Variables:
1.  **Copy the template:** In your project directory, make a copy of the `.env.template` file and rename it to `.env`.
    ```bash
    cp .env.template .env
    ```
2.  **Edit `.env`:** Open the newly created `.env` file with a text editor. You **must** fill in the following critical variables with your specific details:
    *   `EMAIL_ADDRESS`: The target mailbox address to (e.g., `mobile_surveys@cdc.gov`).
    *   `TENANT_ID`: Your Azure AD Tenant ID.
    *   `CLIENT_ID`: The Application (client) ID of your Azure AD App Registration.
    *   `CLIENT_SECRET`: The client secret value for your Azure AD App Registration.
    *   `TARGET_SUBJECT` (Optional, Look for new emails with this subject line. Defaults to "Mobile Phone Usage Query")
    *   `SEARCH_QUESTION` (Optional, defaults to "Do you still need the use of this mobile phone?"): The exact question to find in email bodies.
    *   `OUTPUT_CSV_FILE` (Optional, defaults to "mobile_phone_survey_results.csv"): The name of the CSV file to be generated.
    *   `PROCESSED_FOLDER_NAME` (Optional, defaults to "ProcessedSurveyEmails"): The name of the mailbox's subfolder to move processed emails to.

Install Dependencies: Open your terminal or command prompt, navigate to the project directory, and install the required Python packages:

```bash
pip install -r requirements.txt
```

Run the Script: Execute the Python script from your terminal:

```bash
python email_scanner.py
```

The script will print progress messages to the console. Once finished, you'll find a CSV file (e.g., /output/**mobile_phone_survey_results.csv**, or whatever you named it in .env) in the same directory containing the extracted data.

## Important Considerations

*   **Azure AD App Registration & Permissions:**
    *   You must create an App Registration in Azure Active Directory.
    *   The application needs `Mail.ReadWrite` permission from **Microsoft Graph** (Application permission). This allows the script to read emails from the target mailbox and move them to a processed folder. 
    *   **Admin consent** must be granted for this permission in Azure AD.
    *   A **client secret** must be generated for the app registration (this is your `CLIENT_SECRET`).
    *   Note down the **Application (client) ID** (`CLIENT_ID`) and **Directory (tenant) ID** (`TENANT_ID`).
*   **Application Access Policy (Recommended for Scoped Access):**
    *   With `Mail.ReadWrite` Application permission granted in Azure AD, the application can, by default, access all mailboxes in the organization.
    *   The application's access should be restricted to **only** the specified target mailbox (`EMAIL_ADDRESS`) and adhere to the principle of least privilege, ask IT support to configure an Application Access Policy in Exchange Online.
*   **`.env` File Security:** The `.env` file contains sensitive credentials. The provided `.gitignore` file correctly excludes `.env` from being committed to version control. **Never commit your `.env` file.**
*   **Email Folder:** The script defaults to searching the `INBOX`. If target emails are in a different folder (e.g., "Archive" or a custom folder), ensure the script is configured to target that specific folder. The method for specifying this folder will depend on how the script interacts with the mailbox using the Microsoft Graph API (e.g., by using the folder's ID or well-known name).
*   **Answer Extraction Logic:** The `extract_answer` function uses a simple heuristic (looking for "Yes" or "No" within 100 characters after the question). This might need adjustment based on the exact format of your emails.
* Dependencies: Ensure you have Python and pip installed to manage the packages listed in requirements.txt.
