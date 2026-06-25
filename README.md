# Telegram Reminder Bot

Automated reminder bot for Telegram with Google Sheets integration and GitHub Actions CI/CD.

## Features

- **Period Reminder** — Sends scheduled reminders every Sunday, Tuesday, and Friday at 13:00 UTC with character sheet statistics (Undead Guard, Myrkull Apostates, Scrap) from a Google Sheets spreadsheet.
- **Change Detection** — Monitors the Google Sheet hourly for changes and alerts when updates are detected.
- **Error Alerts** — Notifies on failures with a 1-hour cooldown to avoid spam.
- **Custom Footer** — Configurable footer text via `ganancia.txt`.

## Architecture

```
GitHub Actions (cron) → Google Sheets (XLSX export) → Python (openpyxl) → Telegram Bot API
```

## Workflows

| Workflow | Schedule | Description |
|----------|----------|-------------|
| `send_reminder.yml` | Sun, Tue, Fri @ 13:00 UTC | Sends period reminder with stats |
| `check_changes.yml` | Every hour | Detects sheet modifications |

## Tech Stack

- **Language:** Python 3.12
- **Libraries:** openpyxl
- **CI/CD:** GitHub Actions
- **Integrations:** Telegram Bot API, Google Sheets

## Getting Started

1. Clone the repository
2. Install dependencies: `pip install openpyxl`
3. Set GitHub Secrets:
   - `TELEGRAM_TOKEN` — Your Telegram bot token
   - `TELEGRAM_CHAT_ID` — Target chat ID
4. Push to `main` to trigger the workflows

## Secrets

| Secret | Description |
|--------|-------------|
| `TELEGRAM_TOKEN` | Bot token from BotFather |
| `TELEGRAM_CHAT_ID` | Chat ID to receive messages |
