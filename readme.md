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
    *   `EMAIL_ADDRESS`: The target mailbox the Service Principal has access to (e.g., `your_target_mailbox@example.com`).
    *   `TENANT_ID`: Your Azure AD Tenant ID.
    *   `CLIENT_ID`: The Application (client) ID of your Azure AD App Registration.
    *   `CLIENT_SECRET`: The client secret value for your Azure AD App Registration.
    *   `TARGET_SUBJECT` (Optional, defaults to "Mobile Phone Usage Query"): The subject line the script will search for.
    *   `SEARCH_QUESTION` (Optional, defaults to "Do you still need the use of this mobile phone?"): The exact question to find in email bodies.
    *   `OUTPUT_CSV_FILE` (Optional, defaults to "mobile_phone_survey_results.csv"): The name of the CSV file to be generated.

Install Dependencies: Open your terminal or command prompt, navigate to the project directory, and install the required Python packages:

```bash
pip install -r requirements.txt
```

Run the Script: Execute the Python script from your terminal:

```bash
python email_scanner.py
```

The script will print progress messages to the console. Once finished, you'll find a CSV file (e.g., **mobile_phone_survey_results.csv**, or whatever you named it in .env) in the same directory containing the extracted data.

## Important Considerations

*   **Azure AD App Registration & Permissions:**
    *   You must create an App Registration in Azure Active Directory.
    *   The application needs the `IMAP.AccessAsApp` permission from **Microsoft Graph** (Application permission).
    *   **Admin consent** must be granted for this permission in Azure AD.
    *   A **client secret** must be generated for the app registration (this is your `CLIENT_SECRET`).
    *   Note down the **Application (client) ID** (`CLIENT_ID`) and **Directory (tenant) ID** (`TENANT_ID`).
*   **Mailbox Permissions for Service Principal:** The service principal associated with your Azure AD App Registration needs explicit permission (e.g., `FullAccess`) to the target mailbox (`EMAIL_ADDRESS`). This is typically done via Exchange Online PowerShell. Ask your IT support staff if you need assistance.
*   **`.env` File Security:** The `.env` file contains sensitive credentials. The provided `.gitignore` file correctly excludes `.env` from being committed to version control. **Never commit your `.env` file.**
*   **IMAP Server:** The script defaults to `outlook.office365.com`. This is standard for Microsoft 365, but verify if your environment uses a different server.
*   **Email Folder:** The script defaults to searching the `INBOX`. If target emails are in a different folder (e.g., "Archive" or a custom folder), you'll need to modify the `mailbox.folder.set('YourSpecificFolderName')` line in `email_scanner.py`.
*   **Answer Extraction Logic:** The `extract_answer` function uses a simple heuristic (looking for "Yes" or "No" within 100 characters after the question). This might need adjustment based on the exact format of your emails.
* Dependencies: Ensure you have Python and pip installed to manage the packages listed in requirements.txt.