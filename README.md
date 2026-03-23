# Leave Reply Bot

Automates replying to leave request emails via Outlook web + Telegram approval.

## How it works

1. Logs into Outlook as `Gideon.Buba@nrs.gov.ng`
2. Opens the `leaveservices@nrs.gov.ng` shared mailbox
3. Fetches all unread emails
4. Groq AI classifies each email and drafts a reply
5. Sends draft to your Telegram for approval
6. On `/approve`, sends the reply via Playwright
7. Polls every 5 minutes for new emails going forward

## Setup on your VPS

### 1. Clone / upload the project
```bash
cd ~
# Upload or git clone the project folder here
cd leave-reply-bot
```

### 2. Install dependencies
```bash
npm install
npx playwright install chromium
npx playwright install-deps chromium
```

### 3. Configure environment
```bash
cp .env.example .env
nano .env
```

Fill in:
- `OUTLOOK_EMAIL` — your work email (Gideon.Buba@nrs.gov.ng)
- `OUTLOOK_PASSWORD` — your Outlook password
- `SHARED_MAILBOX` — leaveservices@nrs.gov.ng
- `GROQ_API_KEY` — from your Groq account (same as YouTube pipeline)
- `TELEGRAM_BOT_TOKEN` — your Telegram bot token
- `TELEGRAM_CHAT_ID` — your personal Telegram chat ID

### 4. Get your Telegram Chat ID
Message your bot, then visit:
```
https://api.telegram.org/bot<YOUR_BOT_TOKEN>/getUpdates
```
Find `"id"` under `"chat"` in the response.

### 5. Build and run
```bash
npm run build
npm start
```

### 6. Run with PM2 (keep alive on VPS)
```bash
npm install -g pm2
pm2 start dist/index.js --name leave-bot
pm2 save
pm2 startup
```

## Telegram commands

| Command | Action |
|---|---|
| `/approve` | Send the drafted reply as-is |
| `/edit <new text>` | Replace the draft with your text and send |
| `/skip` | Skip this email, move to next |
| `/type approved` | Redraft as an approval reply |
| `/type denied` | Redraft as a denial reply |
| `/type more_info` | Redraft requesting more information |
| `/type acknowledgement` | Redraft as acknowledgement only |

## Notes

- Session is saved to `session.json` after first login — subsequent runs won't ask for password again
- Delete `session.json` if you need to force a fresh login
- Playwright runs headless (no visible browser window)
- Groq free tier is more than enough for 115+ emails
