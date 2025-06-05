## How to Set Up and Run the Script
Create the Directory: Make a new folder for your project (e.g., outlook_scanner).

Save the Files: Place the .env, email_scanner.py, and requirements.txt files into this directory. Remember to customize .env with your actual Outlook details and preferences.

Install Dependencies: Open your terminal or command prompt, navigate to the project directory, and install the required Python packages:

```bash
cd path/to/your/project
pip install -r requirements.txt
```
Run the Script: Execute the Python script from your terminal:

```bash
python email_scanner.py
```
The script will print progress messages to the console. Once finished, you'll find a CSV file (e.g., **mobile_phone_survey_results.csv**, or whatever you named it in .env) in the same directory containing the extracted data.

## Important Considerations 
* Mailbox Permissions for Service Principal: The service principal associated with your Azure AD App Registration needs explicit permission to access the target mailbox.  Ask your IT support staff to grant the necessary permissions.
* .env File Security:  The .env file contains sensitive credentials (Client ID, Client Secret, Tenant ID). The .gitignore file should exclude the .env file. 
* IMAP Server: The script defaults to outlook.office365.com. This is standard for Microsoft 365, but verify if your environment uses a different server. 
* Email Folder: The script defaults to searching the INBOX. If target emails are in a different folder (e.g., "Archive" or a custom folder), you'll need to modify the mailbox.folder.set('YourSpecificFolderName') line in email_scanner.py. 
* Answer Extraction Logic: The extract_answer function uses a simple heuristic (looking for "Yes" or "No" within 100 characters after the question). This might need adjustment based on the exact format of your emails. 
* Dependencies: Ensure you have Python and pip installed to manage the packages listed in requirements.txt.