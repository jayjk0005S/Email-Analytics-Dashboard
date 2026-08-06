NOTE:

1. Make sure you are using Classic Outlook.
   (If you see the "Try the new Outlook" toggle in the top-right corner,
   turn it OFF.)

2. Open Classic Outlook before running the application.

3. Verify that:
   - Outlook is signed in.
   - Your Inbox is visible.
   - You can see your emails.

Then open Command Prompt and run:

----------------------------------------------------------------------------------------------------------------------

cd "C:\Path\To\Email-Analytics-Dashboard"

py -m venv .venv

.\.venv\Scripts\python.exe -m pip install -r requirements.txt

.\.venv\Scripts\python.exe -m email_analytics.startup

PowerShell -ExecutionPolicy Bypass -File .\scripts\setup_startup_tasks.ps1

-----------------------------------------------------------------------------------------------------------------

Why Classic Outlook?

This application communicates with Outlook using Windows COM (Component Object Model).
The New Outlook does not expose the same COM automation interface, so the listener
works only with Classic Outlook Desktop.
