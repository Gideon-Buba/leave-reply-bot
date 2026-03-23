import * as http from "http";
import * as fs from "fs";
import * as readline from "readline";

const TOKEN_FILE = "./tokens.json";

const CLIENT_ID = process.env.AZURE_CLIENT_ID!;
const TENANT_ID = process.env.AZURE_TENANT_ID!;
const REDIRECT_URI = "http://localhost:3000";
const SCOPES = "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send offline_access";

export interface Email {
  id: string;
  subject: string;
  sender: string;
  senderEmail: string;
  body: string;
  receivedAt: string;
}

interface TokenData {
  access_token: string;
  refresh_token: string;
  expires_at: number;
}

export class OutlookGraph {
  private tokenData: TokenData | null = null;

  async init(): Promise<void> {
    if (fs.existsSync(TOKEN_FILE)) {
      this.tokenData = JSON.parse(fs.readFileSync(TOKEN_FILE, "utf-8"));
      console.log("📂 Loaded saved tokens");
    }
  }

  private saveTokens(data: TokenData): void {
    fs.writeFileSync(TOKEN_FILE, JSON.stringify(data, null, 2));
    this.tokenData = data;
  }

  async login(): Promise<void> {
    if (this.tokenData && Date.now() < this.tokenData.expires_at - 60000) {
      console.log("✅ Already authenticated");
      return;
    }

    if (this.tokenData?.refresh_token) {
      try {
        await this.refreshToken();
        console.log("✅ Token refreshed");
        return;
      } catch {
        console.log("⚠️ Refresh failed, re-authenticating...");
      }
    }

    await this.doOAuthFlow();
  }

  private async doOAuthFlow(): Promise<void> {
    const authUrl =
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?` +
      `client_id=${CLIENT_ID}` +
      `&response_type=code` +
      `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}` +
      `&scope=${encodeURIComponent(SCOPES)}` +
      `&response_mode=query` +
      `&prompt=select_account`;

    console.log("\n🔐 Open this URL in your browser and log in as leaveservices@nrs.gov.ng:\n");
    console.log(authUrl);
    console.log("\nWaiting for authorization...\n");

    const code = await this.waitForCode();
    await this.exchangeCode(code);
    console.log("✅ Authenticated successfully");
  }

  private waitForCode(): Promise<string> {
    return new Promise((resolve, reject) => {
      const server = http.createServer((req, res) => {
        const url = new URL(req.url!, `http://localhost:3000`);
        const code = url.searchParams.get("code");
        const error = url.searchParams.get("error");

        if (error) {
          res.end(`<h2>Error: ${error}</h2>`);
          server.close();
          reject(new Error(error));
          return;
        }

        if (code) {
          res.end(`<h2>✅ Authorization successful! You can close this tab.</h2>`);
          server.close();
          resolve(code);
        }
      });

      server.listen(3000, () => {
        console.log("🌐 Listening on http://localhost:3000 for callback...");
      });

      server.on("error", reject);
    });
  }

  private async exchangeCode(code: string): Promise<void> {
    const body = new URLSearchParams({
      client_id: CLIENT_ID,
      grant_type: "authorization_code",
      code,
      redirect_uri: REDIRECT_URI,
      scope: SCOPES,
    });

    const res = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
      }
    );

    const data = await res.json() as any;
    if (data.error) throw new Error(`Token exchange failed: ${data.error_description}`);

    this.saveTokens({
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      expires_at: Date.now() + data.expires_in * 1000,
    });
  }

  private async refreshToken(): Promise<void> {
    const body = new URLSearchParams({
      client_id: CLIENT_ID,
      grant_type: "refresh_token",
      refresh_token: this.tokenData!.refresh_token,
      scope: SCOPES,
    });

    const res = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
      }
    );

    const data = await res.json() as any;
    if (data.error) throw new Error(data.error_description);

    this.saveTokens({
      access_token: data.access_token,
      refresh_token: data.refresh_token || this.tokenData!.refresh_token,
      expires_at: Date.now() + data.expires_in * 1000,
    });
  }

  private async graphRequest(path: string, options: RequestInit = {}): Promise<any> {
    // Refresh token if needed
    if (this.tokenData && Date.now() > this.tokenData.expires_at - 60000) {
      await this.refreshToken();
    }

    const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${this.tokenData!.access_token}`,
        "Content-Type": "application/json",
        ...options.headers,
      },
    });

    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Graph API error ${res.status}: ${err}`);
    }

    if (res.status === 204) return null;
    return res.json();
  }

  async getUnreadEmails(): Promise<Email[]> {
    console.log("📥 Fetching unread emails...");

    const data = await this.graphRequest(
      "/me/mailFolders/inbox/messages?$filter=isRead eq false&$top=50&$orderby=receivedDateTime asc&$select=id,subject,from,body,receivedDateTime"
    );

    const emails: Email[] = (data.value || []).map((msg: any) => ({
      id: msg.id,
      subject: msg.subject || "No Subject",
      sender: msg.from?.emailAddress?.name || "Unknown",
      senderEmail: msg.from?.emailAddress?.address || "",
      body: msg.body?.content
        ? msg.body.content.replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim().slice(0, 2000)
        : "",
      receivedAt: msg.receivedDateTime || new Date().toISOString(),
    }));

    console.log(`✅ Found ${emails.length} unread emails`);
    return emails;
  }

  async replyToEmail(emailId: string, replyText: string): Promise<boolean> {
    try {
      const replyBody = {
        message: {
          body: {
            contentType: "Text",
            content: replyText,
          },
        },
      };

      await this.graphRequest(`/me/messages/${emailId}/reply`, {
        method: "POST",
        body: JSON.stringify(replyBody),
      });

      // Mark as read
      await this.graphRequest(`/me/messages/${emailId}`, {
        method: "PATCH",
        body: JSON.stringify({ isRead: true }),
      });

      return true;
    } catch (err) {
      console.error("Reply error:", err);
      return false;
    }
  }

  // Keep for compatibility
  async openSharedMailbox(_sharedEmail: string): Promise<void> {
    console.log("✅ Using Graph API — no mailbox switching needed");
  }

  async close(): Promise<void> {
    // Nothing to close
  }
}
