# Email Analytics Dashboard — How It Works

## 1. What this application does

This is a local application that watches the Inbox in **classic Outlook Desktop**, saves selected email information into a local SQLite database, and shows charts and email records in a Streamlit dashboard.

The active flow is:

```text
Classic Outlook Inbox
        ↓
Python Outlook listener
        ↓
Local SQLite database
        ↓
Streamlit dashboard in the browser
```

## 2. What starts when Windows starts

1. You sign in to Windows.
2. The **Email Analytics Dashboard** shortcut in your Windows Startup folder runs automatically.
3. It starts the local dashboard at `http://127.0.0.1:8501`.
4. It starts the Outlook listener in the background.
5. The browser opens the dashboard.

The application runs only on this computer. The dashboard is not public on the internet.

## 3. Outlook connection

The listener connects to **classic Outlook Desktop** already signed in with your mailbox.

It does not use Microsoft Graph or an Entra application registration in the current setup. It uses Outlook's local Windows connection instead.

Important: classic Outlook must be available and signed in. New Outlook does not provide this local connection method.

## 4. First-time import and later catch-up

On the first successful run, the application reads up to the previous 90 days of Inbox emails.

On later runs, it starts from the last successful sync time and checks a small 10-minute overlap. The overlap protects against an email arriving exactly while the computer or listener is restarting.

If the computer is off, emails continue to arrive in Outlook/Microsoft 365 as normal. When you sign in again, the listener catches up on missed emails.

## 5. What is taken from each email

By default, the local database stores metadata only:

- unique email ID
- sender name and email address
- subject
- received date and time
- attachment flag
- importance level

The email body preview is disabled by default, so the body content is not stored. Attachments themselves are not copied into the database.

## 6. Why duplicate emails are avoided

Each mail is saved with a stable identifier. For Outlook Desktop, the application prefers the Internet Message ID and uses the Outlook Entry ID as a fallback.

Before saving, it uses an **upsert** operation:

- a new ID creates a row
- an existing ID updates the same row

So the 10-minute catch-up overlap does not create duplicate dashboard entries.

## 7. What happens for a new email

1. A new email arrives in classic Outlook Inbox.
2. Outlook sends an `ItemAdd` event to the running Python listener.
3. The listener immediately converts the email metadata into a database record.
4. The record is inserted or updated in the local SQLite database.
5. The database commit updates a refresh signal, and open Streamlit sessions refresh in about one second. A five-minute backup refresh recovers a missed signal.
6. The new email becomes visible in the table and charts.

## 8. Local database

The database file is stored here:

```text
data/email_analytics.db
```

inside the Email Analytics Dashboard project folder.

It contains an `emails` table for email records and an `app_state` table for items such as the last successful sync time.

## 9. Dashboard behavior

The dashboard reads only from the local SQLite database.

- **From** filters by sender.
- **Date range** filters by received date.
- **Total Emails** changes to match the chosen filters.
- Charts and the Email Details table use the same selected filters.
- The dashboard refreshes after committed email changes, every five minutes as a backup, and whenever Refresh now is selected.

## 10. Privacy and safety

- Your Outlook password is not stored in this project.
- The dashboard uses the Outlook session already on this Windows computer.
- The database stays local to this computer.
- The dashboard is opened on `127.0.0.1`, which means it is accessible only from this computer.
- No email body is stored unless `STORE_BODY_PREVIEW=true` is explicitly enabled later.

## 11. If something is not running

1. Confirm that classic Outlook is open and signed in.
2. Open `http://127.0.0.1:8501` in a browser.
3. If the dashboard is not open after sign-in, run the Startup shortcut again or start it from the project folder.
4. A refresh or application restart does not delete existing database records.

## 12. Current application components

```text
app.py
    Dashboard interface, filters, charts, table, and auto-refresh.

email_analytics/outlook_listener.py
    Watches classic Outlook for new Inbox emails.

email_analytics/outlook_desktop.py
    Reads Outlook email metadata and performs catch-up syncs.

email_analytics/database.py
    Creates, queries, and upserts data in the local SQLite database.

email_analytics/startup.py
    Starts the listener and dashboard after Windows sign-in.

data/email_analytics.db
    Local database containing the email analytics records.
```
