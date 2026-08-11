# Commands to run

Open Classic Outlook, sign in, and confirm that the Inbox is working. Then open Command Prompt and run:

```cmd
git clone "https://github.com/jayjk0005S/Email-Analytics-Dashboard.git"
cd /d "%USERPROFILE%\Documents\Email-Analytics-Dashboard"
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m email_analytics.startup
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\setup_startup_tasks.ps1"
```

The last command creates the current user's limited-permission Windows logon task. It only needs to be run once.
