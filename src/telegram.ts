import TelegramBot from "node-telegram-bot-api";
import { Email } from "./outlook";
import { ReplyType } from "./groq";

export interface PendingApproval {
  email: Email;
  emailIndex: number;
  draftReply: string;
  suggestedType: ReplyType;
  resolve: (action: ApprovalAction) => void;
}

export type ApprovalAction =
  | { type: "approve" }
  | { type: "edit"; newText: string }
  | { type: "skip" }
  | { type: "change_type"; replyType: ReplyType };

export class TelegramApproval {
  private bot: TelegramBot;
  private chatId: string;
  private pendingApproval: PendingApproval | null = null;

  constructor(token: string, chatId: string) {
    this.bot = new TelegramBot(token, { polling: true });
    this.chatId = chatId;
    this.setupHandlers();
    console.log("🤖 Telegram bot started");
  }

  private setupHandlers(): void {
    this.bot.on("message", async (msg) => {
      const chatId = msg.chat.id.toString();
      if (chatId !== this.chatId) return;

      const text = msg.text?.trim() || "";

      if (!this.pendingApproval) {
        await this.bot.sendMessage(
          this.chatId,
          "No pending email approvals right now.",
        );
        return;
      }

      const { resolve } = this.pendingApproval;

      if (text === "/approve" || text.toLowerCase() === "approve") {
        resolve({ type: "approve" });
        this.pendingApproval = null;
      } else if (text.startsWith("/edit ")) {
        const newText = text.slice(6).trim();
        if (!newText) {
          await this.bot.sendMessage(
            this.chatId,
            "⚠️ Please provide the new reply text after /edit",
          );
          return;
        }
        resolve({ type: "edit", newText });
        this.pendingApproval = null;
      } else if (text === "/skip" || text.toLowerCase() === "skip") {
        resolve({ type: "skip" });
        this.pendingApproval = null;
      } else if (text === "/approve_all") {
        resolve({ type: "approve" });
        this.pendingApproval = null;
      } else if (text.startsWith("/type ")) {
        const replyType = text.slice(6).trim() as ReplyType;
        const validTypes: ReplyType[] = [
          "approved",
          "denied",
          "more_info",
          "acknowledgement",
        ];
        if (validTypes.includes(replyType)) {
          resolve({ type: "change_type", replyType });
          this.pendingApproval = null;
        } else {
          await this.bot.sendMessage(
            this.chatId,
            "⚠️ Valid types: approved, denied, more_info, acknowledgement",
          );
        }
      } else {
        await this.bot.sendMessage(
          this.chatId,
          "Commands:\n/approve — send this draft\n/edit &lt;new text&gt; — replace draft and send\n/skip — skip this email\n/type &lt;type&gt; — redraft with different type",
        );
      }
    });
  }

  async sendForApproval(
    email: Email,
    emailIndex: number,
    draftReply: string,
    suggestedType: ReplyType,
    current: number,
    total: number,
  ): Promise<ApprovalAction> {
    return new Promise(async (resolve) => {
      this.pendingApproval = {
        email,
        emailIndex,
        draftReply,
        suggestedType,
        resolve,
      };

      const typeEmoji: Record<ReplyType, string> = {
        approved: "✅",
        denied: "❌",
        more_info: "❓",
        acknowledgement: "📋",
      };

      const message = [
        `📧 <b>Email ${current}/${total}</b>`,
        ``,
        `<b>From:</b> ${this.esc(email.sender)} &lt;${this.esc(email.senderEmail)}&gt;`,
        `<b>Subject:</b> ${this.esc(email.subject)}`,
        `<b>Received:</b> ${this.esc(email.receivedAt)}`,
        ``,
        `<b>Message:</b>`,
        `<i>${this.esc(email.body.slice(0, 400))}${email.body.length > 400 ? "..." : ""}</i>`,
        ``,
        `━━━━━━━━━━━━━━━━━`,
        `${typeEmoji[suggestedType]} <b>Suggested reply (${suggestedType}):</b>`,
        ``,
        this.esc(draftReply),
        `━━━━━━━━━━━━━━━━━`,
        ``,
        `/approve — send this`,
        `/edit &lt;new text&gt; — replace &amp; send`,
        `/skip — skip this email`,
        `/type approved|denied|more_info|acknowledgement`,
      ].join("\n");

      await this.bot.sendMessage(this.chatId, message, { parse_mode: "HTML" });
    });
  }

  async sendMessage(text: string): Promise<void> {
    await this.bot.sendMessage(this.chatId, text, { parse_mode: "HTML" });
  }

  async sendSummary(
    processed: number,
    approved: number,
    skipped: number,
  ): Promise<void> {
    await this.bot.sendMessage(
      this.chatId,
      `✅ <b>Done!</b>\n\nProcessed: ${processed} emails\nReplied: ${approved}\nSkipped: ${skipped}`,
      { parse_mode: "HTML" },
    );
  }

  private esc(text: string): string {
    return text
      .slice(0, 500)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  stop(): void {
    this.bot.stopPolling();
  }
}
