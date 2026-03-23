import { Browser, BrowserContext, Page, chromium } from "playwright";
import * as fs from "fs";

const SESSION_FILE = "./session.json";

export interface Email {
  id: string;
  subject: string;
  sender: string;
  senderEmail: string;
  body: string;
  receivedAt: string;
}

export class OutlookAutomation {
  private browser: Browser | null = null;
  private context: BrowserContext | null = null;
  private page: Page | null = null;

  async init(): Promise<void> {
    this.browser = await chromium.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });

    // Reuse saved session if available
    if (fs.existsSync(SESSION_FILE)) {
      console.log("📂 Reusing saved session...");
      this.context = await this.browser.newContext({
        storageState: SESSION_FILE,
      });
    } else {
      this.context = await this.browser.newContext();
    }

    this.page = await this.context.newPage();
  }

  async login(email: string, password: string): Promise<void> {
    if (!this.page) throw new Error("Browser not initialized");

    console.log("🔐 Logging into Outlook...");
    await this.page.goto("https://outlook.office.com/mail/");
    await this.page.waitForLoadState("networkidle");

    // Check if already logged in
    const url = this.page.url();
    if (url.includes("outlook.office.com/mail")) {
      console.log("✅ Already logged in via saved session");
      return;
    }

    // Enter email
    await this.page.fill('input[type="email"]', email);
    await this.page.click('input[type="submit"]');
    await this.page.waitForTimeout(1500);

    // Enter password
    await this.page.fill('input[type="password"]', password);
    await this.page.click('input[type="submit"]');
    await this.page.waitForTimeout(2000);

    // Handle "Stay signed in?" prompt
    try {
      await this.page.waitForSelector('input[type="submit"]', {
        timeout: 5000,
      });
      await this.page.click('input[type="submit"]');
    } catch {
      // No prompt, continue
    }

    await this.page.waitForURL("**/outlook.office.com/mail/**", {
      timeout: 15000,
    });

    // Save session for next run
    await this.context!.storageState({ path: SESSION_FILE });
    console.log("✅ Logged in and session saved");
  }

  async openSharedMailbox(sharedEmail: string): Promise<void> {
    if (!this.page) throw new Error("Browser not initialized");

    console.log(`📬 Opening shared mailbox: ${sharedEmail}`);

    // Click profile picture (top right)
    await this.page.click('[aria-label="Account manager"]');
    await this.page.waitForTimeout(1000);

    // Click "Open another mailbox"
    await this.page.click('text=Open another mailbox');
    await this.page.waitForTimeout(1000);

    // Type shared mailbox address
    const input = await this.page.waitForSelector(
      'input[placeholder*="mailbox"], input[type="text"]',
      { timeout: 5000 }
    );
    await input.fill(sharedEmail);
    await this.page.waitForTimeout(500);

    // Click Open / Continue
    await this.page.click('button:has-text("Open"), button:has-text("Continue")');

    // Wait for new tab or navigation
    await this.page.waitForTimeout(3000);

    // Switch to new tab if opened
    const pages = this.context!.pages();
    if (pages.length > 1) {
      this.page = pages[pages.length - 1];
      await this.page.waitForLoadState("networkidle");
    }

    console.log("✅ Shared mailbox opened");
  }

  async getUnreadEmails(): Promise<Email[]> {
    if (!this.page) throw new Error("Browser not initialized");

    console.log("📥 Fetching unread emails...");
    const emails: Email[] = [];

    try {
      // Navigate to inbox
      await this.page.goto(this.page.url().split("/mail")[0] + "/mail/inbox");
      await this.page.waitForLoadState("networkidle");
      await this.page.waitForTimeout(2000);

      // Filter to unread only by clicking Unread filter if available
      try {
        await this.page.click('text=Unread', { timeout: 3000 });
        await this.page.waitForTimeout(1500);
      } catch {
        // No unread filter button, continue
      }

      // Get all email rows
      const emailItems = await this.page.$$('[role="option"], [data-convid]');
      console.log(`Found ${emailItems.length} email items`);

      for (let i = 0; i < Math.min(emailItems.length, 120); i++) {
        try {
          const items = await this.page.$$(
            '[role="option"][aria-selected], [data-convid]'
          );
          if (!items[i]) continue;

          // Check if unread (bold subject or unread indicator)
          const isUnread = await items[i].$('[class*="unread"], [class*="Unread"]');
          if (!isUnread && i > 0) continue;

          await items[i].click();
          await this.page.waitForTimeout(1500);

          // Extract email details
          const subject = await this.page
            .$eval(
              '[data-testid="subject"], [class*="subject"]',
              (el) => el.textContent?.trim() || "No Subject"
            )
            .catch(() => "No Subject");

          const senderName = await this.page
            .$eval(
              '[data-testid="SenderName"], [class*="senderName"], [class*="sender"]',
              (el) => el.textContent?.trim() || "Unknown"
            )
            .catch(() => "Unknown");

          const senderEmail = await this.page
            .$eval(
              '[class*="senderEmail"], [title*="@"]',
              (el) =>
                el.getAttribute("title") || el.textContent?.trim() || ""
            )
            .catch(() => "");

          const body = await this.page
            .$eval(
              '[role="main"] [class*="body"], [data-testid="emailBody"], .ReadMsgBody',
              (el) => el.textContent?.trim() || ""
            )
            .catch(() => "");

          const receivedAt = await this.page
            .$eval(
              '[data-testid="ReceivedTime"], [class*="receivedTime"]',
              (el) => el.textContent?.trim() || new Date().toISOString()
            )
            .catch(() => new Date().toISOString());

          if (subject || body) {
            emails.push({
              id: `email_${i}_${Date.now()}`,
              subject,
              sender: senderName,
              senderEmail,
              body: body.slice(0, 2000), // Cap body length
              receivedAt,
            });
          }
        } catch (err) {
          console.error(`Error extracting email ${i}:`, err);
        }
      }
    } catch (err) {
      console.error("Error fetching emails:", err);
    }

    console.log(`✅ Extracted ${emails.length} unread emails`);
    return emails;
  }

  async replyToEmail(emailIndex: number, replyText: string): Promise<boolean> {
    if (!this.page) throw new Error("Browser not initialized");

    try {
      // Click on the email
      const items = await this.page.$$('[role="option"], [data-convid]');
      if (!items[emailIndex]) {
        console.error(`Email at index ${emailIndex} not found`);
        return false;
      }

      await items[emailIndex].click();
      await this.page.waitForTimeout(1500);

      // Click Reply button
      await this.page.click(
        '[data-testid="reply"], [aria-label="Reply"], button:has-text("Reply")'
      );
      await this.page.waitForTimeout(1500);

      // Find reply compose area and type reply
      const composeArea = await this.page.waitForSelector(
        '[role="textbox"][aria-label*="compose"], [contenteditable="true"][class*="compose"], div[contenteditable="true"]',
        { timeout: 8000 }
      );

      await composeArea.click();
      await this.page.keyboard.press("Control+Home");
      await composeArea.type(replyText + "\n\n", { delay: 10 });

      // Click Send
      await this.page.click(
        '[data-testid="send"], [aria-label="Send"], button:has-text("Send")'
      );
      await this.page.waitForTimeout(2000);

      console.log(`✅ Reply sent for email ${emailIndex}`);
      return true;
    } catch (err) {
      console.error(`Error replying to email ${emailIndex}:`, err);
      return false;
    }
  }

  async close(): Promise<void> {
    if (this.browser) {
      await this.browser.close();
    }
  }
}
