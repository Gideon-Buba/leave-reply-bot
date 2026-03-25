import * as dotenv from "dotenv";
dotenv.config();

import { OutlookGraph } from "./graph";
import { processEmail, draftReply, shouldSkipEmail } from "./groq";
import { TelegramApproval } from "./telegram";

const {
  TELEGRAM_BOT_TOKEN,
  TELEGRAM_CHAT_ID,
  POLL_INTERVAL_MS = "300000",
} = process.env;

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

function assertEnv(): void {
  const required = [
    "GROQ_API_KEY",
    "TELEGRAM_BOT_TOKEN",
    "TELEGRAM_CHAT_ID",
    "AZURE_CLIENT_ID",
    "AZURE_TENANT_ID",
  ];
  const missing = required.filter((k) => !process.env[k]);
  if (missing.length > 0)
    throw new Error(`Missing env vars: ${missing.join(", ")}`);
}

async function processEmails(
  outlook: OutlookGraph,
  telegram: TelegramApproval,
  isFirstRun: boolean,
): Promise<void> {
  const allEmails = await outlook.getUnreadEmails();

  // Filter out reply-chain emails
  const emails = allEmails.filter((e) => !shouldSkipEmail(e));
  const skippedThreads = allEmails.length - emails.length;

  if (emails.length === 0) {
    console.log("📭 No unread emails found");
    if (isFirstRun)
      await telegram.sendMessage(
        "📭 No unread emails in the leave services mailbox.",
      );
    return;
  }

  const threadNote =
    skippedThreads > 0 ? ` (${skippedThreads} reply threads skipped)` : "";
  await telegram.sendMessage(
    `📬 Found *${emails.length}* new email${emails.length > 1 ? "s" : ""}${threadNote}. Starting now...`,
  );

  let approved = 0,
    skipped = 0;

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];
    try {
      // Single API call — classify + draft together
      const { replyType: suggestedType, draft: initialDraft } =
        await processEmail(email);
      let draft = initialDraft;

      let action = await telegram.sendForApproval(
        email,
        i,
        draft,
        suggestedType,
        i + 1,
        emails.length,
      );

      if (action.type === "change_type") {
        await telegram.sendMessage(`🔄 Redrafting as *${action.replyType}*...`);
        draft = await draftReply(email, action.replyType);
        action = await telegram.sendForApproval(
          email,
          i,
          draft,
          action.replyType,
          i + 1,
          emails.length,
        );
      }

      if (action.type === "skip") {
        skipped++;
        await telegram.sendMessage(`⏭️ Skipped.`);
        continue;
      }

      const replyText = action.type === "edit" ? action.newText : draft;
      const sent = await outlook.replyToEmail(email.id, replyText);

      if (sent) {
        approved++;
        await telegram.sendMessage(
          `✅ Reply sent! (${i + 1}/${emails.length})`,
        );
      } else {
        await telegram.sendMessage(
          `⚠️ Failed for "${email.subject}". Skipping.`,
        );
        skipped++;
      }

      // Delay between emails to stay within rate limits
      await sleep(4000);
    } catch (err) {
      console.error(`Error on email ${i}:`, err);
      await telegram.sendMessage(
        `❌ Error on email ${i + 1}: "${email.subject}". Skipping.`,
      );
      skipped++;
      await sleep(4000);
    }
  }

  await telegram.sendSummary(emails.length, approved, skipped);
}

async function main(): Promise<void> {
  assertEnv();
  console.log("🚀 Leave Reply Bot starting...");
  const telegram = new TelegramApproval(TELEGRAM_BOT_TOKEN!, TELEGRAM_CHAT_ID!);
  const outlook = new OutlookGraph();
  await outlook.init();
  await outlook.login();
  await telegram.sendMessage(
    "🚀 *Leave Reply Bot is online!*\n\nFetching emails...",
  );
  await processEmails(outlook, telegram, true);
  console.log(
    `\n⏰ Polling every ${Number(POLL_INTERVAL_MS) / 60000} minutes...`,
  );

  setInterval(async () => {
    try {
      await outlook.login();
      await processEmails(outlook, telegram, false);
    } catch (err) {
      await telegram.sendMessage(
        `⚠️ Poll error: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }, Number(POLL_INTERVAL_MS));
}

main().catch((err) => {
  console.error("Fatal:", err);
  process.exit(1);
});
