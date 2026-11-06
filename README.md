# Telegram Automations

Shrubbery is the first worker. It reads Inbox/Shrubbery, sends concise updates to the configured Telegram chat and writes eligible events to iCloud School. New event titles capitalise every word.

## Setup on a Linux VPS

Install Python 3.11 or newer with venv support, curl and unzip. Use a regular service account with an active systemd user manager. Enable lingering for that account with `sudo loginctl enable-linger "$(id -un)"` so the worker keeps running after logout.

1. Register a Microsoft Entra application that supports personal Microsoft accounts. Enable public client flows and add the delegated Microsoft Graph Mail.Read permission. Copy the Application (client) ID. The worker only reads Inbox/Shrubbery, although Microsoft grants Mail.Read at mailbox scope.
2. Generate an Apple app-specific password for CalDAV and have your OpenAI API key, replacement Telegram token and saved chat ID ready. Start a chat with the replacement bot if you created a different bot.
3. Run the installer from the extracted repository using `bash install.sh`. Secrets are entered without echo and stored locally under the ignored state directory. The Microsoft device code is displayed in the terminal. Enter it on the Microsoft page shown, then complete the Authenticator prompt Microsoft presents. The service caches and refreshes its own tokens. To renew revoked access, stop the service, run `bash install.sh --login`, then restart the service.
4. To resume from the last completed Shrubbery message from the old automation, enter `2026-11-06T15:48:08Z` at the checkpoint prompt. Leave it blank to start at installation time.
5. For push, set up a public HTTPS reverse proxy forwarding POST /graph to http://127.0.0.1:8765/graph and enter its full HTTPS URL during setup. The worker renews the Graph subscription and checks for missed emails every minute. With no reachable webhook, it checks only the folder every 20 seconds, subject to throttling.

Check status with `systemctl --user status telegram-shrubbery` and logs with `journalctl --user -u telegram-shrubbery -f`.

The worker checks School before adding events. Existing entries and date changes that need manual reconciliation remain in School. Review Telegram alerts and the journal for unreadable attachments, timing issues or uncertain delivery. Telegram can accept a request even if the connection drops before acknowledging it. The worker stops processing at that message to prevent a duplicate send.
