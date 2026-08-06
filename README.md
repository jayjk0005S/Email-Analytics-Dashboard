# Email Analytics Dashboard

A private, local Streamlit dashboard for monitoring one Outlook Inbox. The application imports email metadata into SQLite and presents sender activity, daily volume, recent messages, and Inbox trends in a readable dark-mode interface.

The default connection is **Classic Outlook Desktop on Windows**. Microsoft Graph is supported as an optional alternative.

```text
Classic Outlook Desktop -> listener and catch-up sync -> SQLite -> Streamlit dashboard
```

## Dashboard behaviour

The dashboard provides:

- A status bar showing the active mail source, storage mode, refresh interval, last successful sync, and Graph webhook status.
- Sender and date-range filters. The initial date range is **yesterday through today**.
- A **Total Emails** metric that updates with the selected filters.
- **Emails by Sender**, showing the 15 most active senders.
- **Email Details**, showing the latest sender, received time, subject, attachment status, and importance.
- **Email Trends Over Time**, showing daily Inbox activity.
- **Daily Email Volume**, showing the latest 12 days in the filtered data.
- Email-triggered refresh in about one second, with a five-minute backup refresh.
- Received times displayed in the `Asia/Kolkata` time zone.
- A balanced, screen-fitting interface with readable text at normal browser zoom.

## Privacy and local storage

By default, the application stores only the sender, subject, received time, attachment flag, importance, and source identifiers required to prevent duplicates. The SQLite database remains on this computer.

Full message bodies, HTML, and attachment files are not stored. A short plain-text body preview is saved only when `STORE_BODY_PREVIEW=true` is explicitly set in `.env`.

Private files such as `.env`, the SQLite database, and the Graph token cache are ignored by Git.

## Requirements

- Windows 10 or Windows 11
- Python 3.11 or newer
- Classic Outlook Desktop installed and signed in for the default connection mode
- Microsoft Entra application registration only for the optional Graph mode

The Classic Outlook integration uses Windows COM and does not connect to the new Outlook application.

## Installation

Open PowerShell in this project directory:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

The default settings work with Classic Outlook Desktop. Creating a `.env` file is optional unless you need to change a setting:

```powershell
Copy-Item .env.example .env
```

## Run with Classic Outlook Desktop

1. Open Classic Outlook and confirm that the correct mailbox is signed in.
2. Run the initial Inbox import:

```powershell
python -m email_analytics.mail_sync --initial-days 90
```

3. Start the dashboard and live Outlook listener together:

```powershell
python -m email_analytics.startup
```

The startup command performs a duplicate-safe catch-up import, starts the new-mail listener, starts Streamlit on `http://127.0.0.1:8501`, and opens the dashboard in a new default-browser tab. At Windows sign-in it waits for both Streamlit and the interactive desktop, retries the tab launch for up to one minute, and records the result in `data/startup.log`.

Use `--no-browser` when the services should start without opening a new tab:

```powershell
python -m email_analytics.startup --no-browser
```

Each new Classic Outlook Inbox item is saved immediately. Stable Outlook message identifiers and SQLite upserts prevent duplicate records. If the listener was offline, its next startup catch-up import retrieves missed messages.

## Run with demo data

Demo mode is useful when Outlook is unavailable or when previewing the dashboard:

```powershell
python -m email_analytics.seed_demo
streamlit run app.py
```

This creates 511 deterministic demo records using `example.com` addresses. It does not access a real mailbox.

## Start automatically at Windows sign-in

After installation and one successful manual run, create the local scheduled task:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\setup_startup_tasks.ps1
```

In Classic Outlook mode, this creates **Email Analytics - Start Dashboard**, which launches the dashboard and Outlook listener in the signed-in Windows session.

When `MAIL_SOURCE=graph`, the script also creates **Email Analytics - Sync Inbox**, which performs a duplicate-safe catch-up sync every ten minutes.

## Optional Microsoft Graph mode

Graph mode reads the signed-in user's Inbox through delegated `Mail.Read` permission. It never asks for or stores the mailbox password.

1. Create an app registration in the Microsoft Entra admin center.
2. Under **Authentication**, enable the public-client/device-code flow for mobile and desktop applications.
3. Add the Microsoft Graph delegated permission `Mail.Read`. Administrator consent may be required by the organization.
4. Copy `.env.example` to `.env` and configure:

```text
MAIL_SOURCE=graph
GRAPH_CLIENT_ID=your-application-client-id
GRAPH_TENANT_ID=your-directory-tenant-id
```

Use `organizations` as `GRAPH_TENANT_ID` when a tenant-specific ID is not required.

Run the first Graph sync:

```powershell
python -m email_analytics.mail_sync --initial-days 90
```

The terminal displays a Microsoft device sign-in URL and code. After sign-in and consent, the application imports the Inbox for that account. Start the Graph-mode services with:

```powershell
python -m email_analytics.startup
```

## Optional Graph webhook updates

The Graph webhook listener can receive a notification when a new Inbox message is created, fetch that message, and store it immediately:

```text
Graph notification -> webhook listener -> fetch message -> SQLite upsert -> dashboard refresh
```

Microsoft Graph requires a publicly reachable HTTPS endpoint; it cannot send notifications directly to `localhost`. Configure these private `.env` values:

```text
PUBLIC_WEBHOOK_URL=https://your-public-host/webhooks/graph
WEBHOOK_CLIENT_STATE=a-private-random-value-at-least-24-characters
```

Create and inspect the subscription:

```powershell
python -m email_analytics.subscriptions create
python -m email_analytics.subscriptions status
```

Renew it before expiration:

```powershell
python -m email_analytics.subscriptions renew
```

Keep the periodic Graph sync enabled because it catches messages missed while the computer or webhook endpoint is unavailable.

### Test the webhook locally

The local demo endpoint proves the immediate database update and dashboard refresh without a public endpoint:

```powershell
# Terminal 1
python -m email_analytics.webhook

# Terminal 2
Invoke-RestMethod -Method Post -Uri http://127.0.0.1:8787/demo/new-email `
  -ContentType 'application/json' `
  -Body '{"sender_email":"new.sender@example.com","sender_name":"Live Demo","subject":"Email arrived just now"}'
```

The simulated record updates the shared refresh signal and appears in an open dashboard in about one second.

## Project structure

- `app.py` - Streamlit dashboard, filters, table, and Plotly charts.
- `email_analytics/database.py` - SQLite schema, queries, and duplicate-safe upserts.
- `email_analytics/outlook_desktop.py` - Classic Outlook Inbox conversion and catch-up sync.
- `email_analytics/outlook_listener.py` - Immediate Classic Outlook new-mail listener.
- `email_analytics/mail_sync.py` - Selects the configured Outlook or Graph sync implementation.
- `email_analytics/graph_client.py` - Microsoft Graph authentication and Inbox requests.
- `email_analytics/webhook.py` - Graph webhook receiver and local demo endpoint.
- `email_analytics/subscriptions.py` - Graph webhook subscription creation and renewal.
- `email_analytics/startup.py` - Starts the required local services.
- `email_analytics/seed_demo.py` - Generates deterministic demonstration records.
- `scripts/setup_startup_tasks.ps1` - Configures Windows sign-in and Graph catch-up tasks.

## Verification

Install the test runner if needed, then run:

```powershell
python -m pip install pytest
python -m pytest tests -q
```

## Data-safety notes

- Do not upload `.env`, `data/email_analytics.db`, or `data/token_cache.bin` to chat, email, or source control.
- The Graph token cache is protected with Windows DPAPI and can be used only by the same Windows user.
- Demo messages use `example.com` addresses and are not real emails.
