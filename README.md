# Email Analytics Dashboard

Dashboard filters cascade from left to right: date/search narrows the available days, senders,
importance values, attachment states, and read statuses. Calendar filtering uses Asia/Kolkata
dates, and dropdown/calendar popups use the dashboard's light lavender theme.

A local Windows application that reads the Inbox configured in Classic Outlook Desktop, stores email metadata in SQLite, and displays analytics in a Streamlit dashboard.

```text
Classic Outlook Desktop -> Windows COM listener -> SQLite -> Streamlit dashboard
```

Each installation reads the Outlook Inbox of the Windows user running it. Its database stays on that computer.

## Requirements

- Windows 10 or Windows 11
- Python 3.11 or newer
- Git
- Classic Outlook Desktop installed
- The user's mailbox signed in and the Inbox working in Classic Outlook

## Install and run on another computer

Open Classic Outlook first and confirm that the Inbox displays emails. Then open PowerShell in the folder that should contain the project and run:

```powershell
git clone "https://github.com/jayjk0005S/Email-Analytics-Dashboard.git"

cd Email-Analytics-Dashboard

py -m venv .venv

.\.venv\Scripts\python.exe -m pip install -r requirements.txt

.\.venv\Scripts\python.exe -m email_analytics.startup

PowerShell -ExecutionPolicy Bypass -File .\scripts\setup_startup_tasks.ps1
```

The startup command:

1. Connects to the default Inbox in Classic Outlook.
2. Creates `data/email_analytics.db` automatically.
3. Imports recent Inbox messages without creating duplicates.
4. Starts the listener for newly received messages.
5. Starts the Streamlit dashboard.
6. Opens `http://127.0.0.1:8501` in the default browser.

In **Email Details**, select one row to display **Open in Outlook** and **Reply in Outlook**. Reply opens an unsent Outlook draft so the user can review and send it manually.

The final PowerShell command creates the Windows scheduled task **Email Analytics - Start Dashboard**. After that, the application starts automatically whenever that user signs into Windows.

## Run manually later

```powershell
cd "C:\Path\To\Email-Analytics-Dashboard"

.\.venv\Scripts\python.exe -m email_analytics.startup
```

## Update the application

```powershell
cd "C:\Path\To\Email-Analytics-Dashboard"

git pull

.\.venv\Scripts\python.exe -m pip install -r requirements.txt

.\.venv\Scripts\python.exe -m email_analytics.startup
```

## Local data and privacy

The application creates the `data/` directory automatically. It contains the local SQLite database, listener state, and startup logs. The directory is excluded from Git because the database can contain private senders and email subjects.

Do not upload or send the `data/` directory. Each person receives their own Outlook data when they run the application on their computer.

## Main files

- `app.py` - Streamlit dashboard
- `email_analytics/outlook_desktop.py` - reads Outlook Inbox messages
- `email_analytics/outlook_listener.py` - listens for new messages
- `email_analytics/database.py` - stores and queries local email data
- `email_analytics/startup.py` - starts the listener, dashboard, and browser
- `scripts/setup_startup_tasks.ps1` - configures automatic Windows startup

For a shorter copy-and-run guide, see `Commands-to-run.md`.
