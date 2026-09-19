# Gmail To Excel Automation
<img width="1440" height="900" alt="image" src="https://github.com/user-attachments/assets/73cfe7ff-e2da-444b-83a2-4c3c991e3e46" />

<img width="1440" height="900" alt="image" src="https://github.com/user-attachments/assets/adf0d911-262c-428b-a84e-cafafa0cafac" />

<img width="1440" height="900" alt="image" src="https://github.com/user-attachments/assets/06041ccd-caad-4e5f-a90c-6cbc5f736963" />

<img width="1440" height="900" alt="image" src="https://github.com/user-attachments/assets/7f4ec1f3-bd74-4bce-ae71-2448e454ba16" />

A Python-based desktop automation project that connects to Gmail using the **Google Gmail API**, retrieves emails based on user-defined criteria, extracts email information, and exports the results into structured Excel reports.

The project is designed for practical **email data extraction, reporting, analysis, and Excel automation**.

---

## 📌 Project Overview

**Gmail To Excel Automation** automates the following workflow:

```text
Gmail
  │
  ▼
Google Gmail API
  │
  ▼
OAuth 2.0 Authentication
  │
  ├── credentials.json
  │
  └── token.json
  │
  ▼
Email Search & Filtering
  │
  ▼
Email Data Extraction
  │
  ▼
Data Processing
  │
  ▼
Excel Report
```

### Example Use Cases

* Extract Gmail emails into Excel
* Search emails by sender
* Search by subject
* Search by date
* Extract email addresses
* Extract subject and message information
* Extract received date/time
* Filter emails using Gmail search queries
* Create structured Excel reports
* Automate repetitive email reporting
* Analyze email communication data

---

# 🚀 Key Features

## Gmail Integration

* Google Gmail API integration
* OAuth 2.0 authentication
* Gmail message search
* Sender filtering
* Subject filtering
* Date filtering
* Gmail search query support
* Email metadata extraction

## Excel Automation

* Export Gmail data to Excel
* Structured worksheets
* Automatic column creation
* Excel formatting
* Data filtering
* Report generation
* Excel-ready datasets

## Security

* OAuth 2.0 authentication
* No Gmail password stored in the application
* `credentials.json` excluded from GitHub
* `token.json` excluded from GitHub
* `.gitignore` protection
* Local OAuth token storage

---

# 🛠️ Technologies Used

| Technology   | Purpose                 |
| ------------ | ----------------------- |
| Python       | Application development |
| Gmail API    | Gmail integration       |
| Google Cloud | API configuration       |
| OAuth 2.0    | Authentication          |
| Tkinter      | Desktop GUI             |
| Pandas       | Data processing         |
| OpenPyXL     | Excel generation        |
| Git          | Version control         |
| GitHub       | Source-code repository  |

---

# 📁 Project Structure

```text
Gmail_To_Excel_Automation/
│
├── main.py
├── README.md
├── requirements.txt
├── .gitignore
│
├── config/
│   ├── __init__.py
│   └── settings.py
│
├── gmail/
│   ├── __init__.py
│   ├── gmail_auth.py
│   ├── gmail_reader.py
│   ├── email_parser.py
│   └── email_filter.py
│
├── excel/
│   ├── __init__.py
│   ├── excel_writer.py
│   ├── excel_formatter.py
│   └── report_generator.py
│
├── data/
│   ├── input/
│   └── output/
│
├── logs/
│
└── tests/
```

---

# ☁️ Google Cloud Setup

The Gmail API requires a Google Cloud project.

## Step 1 — Open Google Cloud Console

Open:

https://console.cloud.google.com/

Sign in using your Google account.

---

## Step 2 — Create a Google Cloud Project

1. Open **Google Cloud Console**
2. Click the project selector
3. Click **New Project**
4. Enter a project name

Example:

```text
Gmail To Excel Automation
```

5. Click **Create**

---

# Step 3 — Enable Gmail API

Inside your Google Cloud project:

1. Open **APIs & Services**
2. Select **Library**
3. Search for:

```text
Gmail API
```

4. Select **Gmail API**
5. Click **Enable**

The application can now request access to Gmail through Google's API.

---

# Step 4 — Configure OAuth Consent Screen

Open:

```text
Google Cloud Console
→ APIs & Services
→ OAuth consent screen
```

Configure the application.

### Application information

Example:

```text
App name:
Gmail To Excel Automation
```

Enter your developer/contact email information as requested by Google.

---

# Step 5 — Configure Test Users

If the OAuth application is configured for testing, add the Gmail account that will be used with the application as a **Test User**.

Example:

```text
yourname@gmail.com
```

This is important when the OAuth application has not been published for general use.

---

# Step 6 — Configure Gmail API Scopes

The application needs Gmail permissions depending on the functionality implemented.

For read-only email extraction, a commonly used scope is:

```text
https://www.googleapis.com/auth/gmail.readonly
```

This allows the application to read Gmail data without allowing it to modify or delete messages.

Only request the minimum permissions required by the application.

---

# 🔐 OAuth Credentials

The application uses two important local files:

```text
credentials.json
token.json
```

These files have different purposes.

---

# 📄 credentials.json

`credentials.json` contains the OAuth client configuration downloaded from Google Cloud.

It is generated from:

```text
Google Cloud Console
        ↓
APIs & Services
        ↓
Credentials
        ↓
Create Credentials
        ↓
OAuth Client ID
```

For a desktop Tkinter application, choose the appropriate **Desktop application** OAuth client type.

Google provides a JSON file that is commonly downloaded as:

```text
credentials.json
```

Place it in the location expected by your application, for example:

```text
Gmail_To_Excel_Automation/
│
├── credentials.json
├── main.py
└── ...
```

### Important

`credentials.json` contains OAuth client information and must **not be committed to GitHub**.

Add it to `.gitignore`:

```gitignore
credentials.json
```

---

# 🔑 token.json

`token.json` is generated by the application after successful OAuth authentication.

Typical process:

```text
Application starts
       ↓
credentials.json loaded
       ↓
Google OAuth browser window opens
       ↓
User signs in
       ↓
User grants Gmail permission
       ↓
Google returns OAuth tokens
       ↓
token.json created
```

The application can subsequently use the stored token instead of asking the user to authenticate every time, subject to Google's token/session rules.

Example:

```text
Gmail_To_Excel_Automation/
│
├── credentials.json
├── token.json
├── main.py
└── ...
```

### Important

`token.json` may contain OAuth access/refresh credentials.

**Never upload `token.json` to GitHub.**

Add:

```gitignore
token.json
```

to `.gitignore`.

---

# 🚨 Security Warning

Never commit these files:

```text
credentials.json
token.json
client_secret.json
.env
```

Recommended `.gitignore`:

```gitignore
# Python
__pycache__/
*.py[cod]

# Virtual environments
venv/
env/
.venv/

# Google OAuth
credentials.json
token.json
client_secret.json

# Environment variables
.env
.env.*

# Logs
*.log
logs/

# Generated Excel files
*.xlsx
*.xls

# IDE
.vscode/
.idea/

# Windows
Thumbs.db
Desktop.ini
```

---

# ⚠️ If credentials are accidentally uploaded

If `credentials.json` or `token.json` is accidentally pushed to GitHub:

1. Do not continue using the exposed credentials.
2. Revoke/rotate the affected Google credentials.
3. Remove the files from Git history.
4. Add the files to `.gitignore`.
5. Generate fresh credentials/tokens if required.
6. Push the cleaned repository.

Removing a file from the latest working directory is **not necessarily enough** if the secret remains in previous Git commits.

---

# 🐍 Python Environment Setup

## Step 1 — Open PowerShell

```powershell
cd D:\Gmail_To_Excel
```

---

## Step 2 — Create Virtual Environment

```powershell
python -m venv venv
```

---

## Step 3 — Activate Environment

```powershell
.\venv\Scripts\Activate.ps1
```

If PowerShell prevents activation, use:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Then:

```powershell
.\venv\Scripts\Activate.ps1
```

---

# 📦 Install Dependencies

Example:

```powershell
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib pandas openpyxl
```

For Tkinter, use the Python installation that includes Tk support.

---

# 📋 requirements.txt

Example:

```text
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
pandas
openpyxl
```

Install everything using:

```powershell
pip install -r requirements.txt
```

---

# ▶️ Running the Application

From the project directory:

```powershell
cd D:\Gmail_To_Excel
```

Activate the environment:

```powershell
.\venv\Scripts\Activate.ps1
```

Run:

```powershell
python main.py
```

On the first run, the application may open a Google authentication page.

After successful authentication:

```text
token.json
```

may be generated locally.

---

# 📧 Gmail Search

The application can use Gmail search syntax.

Examples:

### Search by sender

```text
from:example@gmail.com
```

### Search by subject

```text
subject:invoice
```

### Search unread messages

```text
is:unread
```

### Search attachments

```text
has:attachment
```

### Search after a date

```text
after:2026/01/01
```

### Combined search

```text
from:example@gmail.com subject:invoice has:attachment
```

The exact supported queries depend on how the application implements the Gmail search interface.

---

# 📊 Excel Output

Example output:

```text
data/
└── output/
    └── Gmail_Email_Report.xlsx
```

Possible columns:

| Column           | Description                  |
| ---------------- | ---------------------------- |
| Email_ID         | Gmail message identifier     |
| Thread_ID        | Gmail conversation/thread ID |
| Sender           | Email sender                 |
| Recipient        | Email recipient              |
| Subject          | Email subject                |
| Date             | Received date                |
| Time             | Received time                |
| Labels           | Gmail labels                 |
| Has_Attachment   | Attachment indicator         |
| Attachment_Count | Number of attachments        |
| Email_Body       | Extracted email content      |

---

# 📈 Example Workflow

```text
Start Application
       │
       ▼
Connect Gmail
       │
       ▼
OAuth Authentication
       │
       ▼
Enter Gmail Search Criteria
       │
       ▼
Search Emails
       │
       ▼
Display Email Results
       │
       ▼
Select Records
       │
       ▼
Extract Email Data
       │
       ▼
Process Data
       │
       ▼
Generate Excel
       │
       ▼
Open Excel Report
```

---

# 🖥️ Planned GUI

The project can provide a user-friendly Tkinter interface with:

```text
+------------------------------------------------------+
|          Gmail To Excel Automation                   |
+------------------------------------------------------+
|                                                      |
|  Gmail Account: [ Connected / Not Connected ]        |
|                                                      |
|  Search Query: [ from:example@gmail.com           ]  |
|                                                      |
|  Date From:   [ DD/MM/YYYY ]                         |
|  Date To:     [ DD/MM/YYYY ]                         |
|                                                      |
|  [ Connect Gmail ]  [ Search Emails ]                |
|                                                      |
|  --------------------------------------------------  |
|  Email Results                                      |
|  --------------------------------------------------  |
|  Sender | Subject | Date | Attachment | Select      |
|                                                      |
|  [ Select All ] [ Clear ]                            |
|                                                      |
|  [ Export To Excel ]                                 |
|                                                      |
|  Status: Ready                                       |
+------------------------------------------------------+
```

---

# 🧪 Testing

Run the application:

```powershell
python main.py
```

Test the following:

* Gmail authentication
* Gmail search
* Email retrieval
* Sender extraction
* Subject extraction
* Date extraction
* Attachment detection
* Excel generation
* Excel formatting
* Error handling

---

# 🔍 Troubleshooting

## `credentials.json not found`

Check that the file exists locally:

```text
D:\Gmail_To_Excel\credentials.json
```

Do not upload it to GitHub.

---

## `token.json` not found

This may be normal on the first run.

The application can create `token.json` after successful OAuth authentication, depending on the implementation.

---

## Access blocked by Google

Check:

```text
Google Cloud Console
→ APIs & Services
→ OAuth consent screen
→ Test users
```

Make sure the Gmail account you're using is authorized as a test user when the OAuth application is in testing mode.

---

## Gmail API has not been used

Check:

```text
Google Cloud Console
→ APIs & Services
→ Library
→ Gmail API
```

Make sure the Gmail API is enabled.

---

## GitHub rejects the push because of a secret

Make sure these are excluded:

```text
credentials.json
token.json
client_secret.json
```

Check tracked files:

```powershell
git ls-files credentials.json
git ls-files token.json
```

These commands should return no output.

---

# 🔒 Recommended Security Architecture

For development:

```text
Local Computer
│
├── credentials.json       ← Local only
├── token.json             ← Local only
│
└── Python Application
          │
          ▼
      Gmail API
```

GitHub:

```text
GitHub Repository
│
├── Python source code
├── README.md
├── requirements.txt
├── .gitignore
│
└── NO OAuth secrets
```

---

# 📌 GitHub Repository

Repository:

**Gmail To Excel Automation**

```text
https://github.com/prabhakardl/Gmail_To_Excel_Automation
```

---

# 🎯 Future Enhancements

Planned functionality can include:

* Gmail folder/label selection
* Advanced search builder
* Multiple Gmail accounts
* Email attachment extraction
* Attachment download
* Duplicate email detection
* HTML email parsing
* Email body cleaning
* Excel dashboard
* Power BI integration
* Scheduled email extraction
* Automatic Excel report generation
* Email analytics
* Sender statistics
* Subject analysis
* Daily/weekly/monthly reports
* Export to CSV
* Export to Excel
* Logging and audit trail
* Error reporting
* Background processing
* Search history
* Configuration management

---

# 👨‍💻 Project Purpose

This project demonstrates practical skills in:

* Python automation
* Gmail API integration
* Google Cloud configuration
* OAuth 2.0
* API-based data extraction
* Data cleaning
* Data transformation
* Excel automation
* Pandas
* OpenPyXL
* Tkinter GUI development
* Git/GitHub
* Secure credential management

---


