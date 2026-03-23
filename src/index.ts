import * as dotenv from "dotenv";
dotenv.config();

import { OutlookAutomation } from "./outlook";
import { draftReply, classifyEmail } from "./groq";
import { TelegramApproval } from "./telegram";

const {
  OUTLOOK_EMAIL,
  OUTLOOK_PASSWORD,
  SHARED_MAILBOX,
  GROQ_API_KEY,
  TELEGRAM_BOT_TOKEN,
  TELEGRAM_CHAT_ID,
  POLL_INTERVAL_MS = "300000",
} = process.env;

function assertEnv(): void {
  const required = [
    "OUTLOOK_EMAIL",
    "OUTLOOK_PASSWORD",
    "SHARED_MAILBOX",
    "GROQ_API_KEY",
    "TELEGRAM_BOT_TOKEN",
    "TELEGRAM_CHAT_ID",
  ];
  const missing = required.filter((k) => !process.env[k]);
  if (missing.length > 0) {
    throw new Error(`Missing env vars: ${missing.join(", ")}`);
  }
}

async function processEmails(
  outlook: OutlookAutomation,
  telegram: TelegramApproval,
  isFirstRun: boolean
): Promise<void> {
  const emails = await outlook.getUnreadEmails();

  if (emails.length === 0) {
    console.log("📭 No unread emails found");
    if (isFirstRun) {
      await telegram.sendMessage(
        "📭 No unread emails found in the leave services mailbox."
      );
    }
    return;
  }

  console.log(`📬 Processing ${emails.length} unread emails...`);
  await telegram.sendMessage(
    `📬 Found *${emails.length}* unread email${emails.length > 1 ? "s" : ""} to process. Starting now...`
  );

  let approved = 0;
  let skipped = 0;

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];
    console.log(`\n[${i + 1}/${emails.length}] Processing: ${email.subject}`);

    try {
      // Classify the email
      const suggestedType = await classifyEmail(email);
      console.log(`  → Suggested type: ${suggestedType}`);

      // Draft reply
      let draft = await draftReply(email, suggestedType);
      console.log(`  → Draft ready`);

      // Send to Telegram for approval
      let action = await telegram.sendForApproval(
        email,
        i,
        draft,
        suggestedType,
        i + 1,
        emails.length
      );

      // Handle redraft request
      if (action.type === "change_type") {
        console.log(`  → Redrafting as: ${action.replyType}`);
        await telegram.sendMessage(
          `🔄 Redrafting as *${action.replyType}*...`
        );
        draft = await draftReply(email, action.replyType);

        // Send redraft for re-approval
        action = await telegram.sendForApproval(
          email,
          i,
          draft,
          action.replyType,
          i + 1,
          emails.length
        );
      }

      if (action.type === "skip") {
        console.log(`  → Skipped`);
        skipped++;
        await telegram.sendMessage(`⏭️ Skipped.`);
        continue;
      }

      const replyText =
        action.type === "edit" ? action.newText : draft;

      // Send the reply via Playwright
      const sent = await outlook.replyToEmail(i, replyText);

      if (sent) {
        approved++;
        await telegram.sendMessage(`✅ Reply sent! (${i + 1}/${emails.length})`);
      } else {
        await telegram.sendMessage(
          `⚠️ Failed to send reply for "${email.subject}". Skipping.`
        );
        skipped++;
      }

      // Small delay between emails to avoid rate limiting
      await new Promise((r) => setTimeout(r, 1500));
    } catch (err) {
      console.error(`Error processing email ${i}:`, err);
      await telegram.sendMessage(
        `❌ Error on email ${i + 1}: "${email.subject}". Skipping.`
      );
      skipped++;
    }
  }

  await telegram.sendSummary(emails.length, approved, skipped);
}

async function main(): Promise<void> {
  assertEnv();

  console.log("🚀 Leave Reply Bot starting...");

  const telegram = new TelegramApproval(
    TELEGRAM_BOT_TOKEN!,
    TELEGRAM_CHAT_ID!
  );

  await telegram.sendMessage(
    "🚀 *Leave Reply Bot is online!*\n\nConnecting to Outlook..."
  );

  const outlook = new OutlookAutomation();

  try {
    await outlook.init();
    await outlook.login(OUTLOOK_EMAIL!, OUTLOOK_PASSWORD!);
    await outlook.openSharedMailbox(SHARED_MAILBOX!);

    await telegram.sendMessage("✅ Connected to Outlook. Fetching emails...");

    // Process existing unread emails on first run
    await processEmails(outlook, telegram, true);

    // Then poll for new emails on interval
    console.log(
      `\n⏰ Polling for new emails every ${Number(POLL_INTERVAL_MS) / 60000} minutes...`
    );

    setInterval(async () => {
      console.log("\n🔄 Checking for new emails...");
      try {
        await processEmails(outlook, telegram, false);
      } catch (err) {
        console.error("Poll error:", err);
        // Reinitialize browser on error
        await outlook.close();
        await outlook.init();
        await outlook.login(OUTLOOK_EMAIL!, OUTLOOK_PASSWORD!);
        await outlook.openSharedMailbox(SHARED_MAILBOX!);
      }
    }, Number(POLL_INTERVAL_MS));
  } catch (err) {
    console.error("Fatal error:", err);
    await telegram.sendMessage(
      `❌ Fatal error: ${err instanceof Error ? err.message : String(err)}`
    );
    await outlook.close();
    telegram.stop();
    process.exit(1);
  }
}

main();
