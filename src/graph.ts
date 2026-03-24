import * as fs from "fs";

const TOKEN_FILE = "./tokens.json";
const CLIENT_ID = process.env.AZURE_CLIENT_ID!;
const TENANT_ID = process.env.AZURE_TENANT_ID!;
const SCOPES =
  "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send offline_access";

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
    await this.doDeviceCodeFlow();
  }

  private async doDeviceCodeFlow(): Promise<void> {
    const res = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/devicecode`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: CLIENT_ID,
          scope: SCOPES,
        }).toString(),
      },
    );
    const data = (await res.json()) as any;
    if (data.error)
      throw new Error(`Device code error: ${data.error_description}`);

    console.log("\n🔐 To authorize, go to:", data.verification_uri);
    console.log("📋 Enter this code:", data.user_code);
    console.log("\nWaiting for you to complete login in the browser...\n");

    await this.pollForToken(data.device_code, data.interval || 5);
    console.log("✅ Authenticated successfully");
  }

  private async pollForToken(
    deviceCode: string,
    interval: number,
  ): Promise<void> {
    while (true) {
      await new Promise((r) => setTimeout(r, interval * 1000));
      const res = await fetch(
        `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: CLIENT_ID,
            grant_type: "urn:ietf:params:oauth:grant-type:device_code",
            device_code: deviceCode,
          }).toString(),
        },
      );
      const data = (await res.json()) as any;
      if (data.access_token) {
        this.saveTokens({
          access_token: data.access_token,
          refresh_token: data.refresh_token,
          expires_at: Date.now() + data.expires_in * 1000,
        });
        return;
      }
      if (data.error === "authorization_pending") continue;
      if (data.error === "slow_down") {
        interval += 5;
        continue;
      }
      throw new Error(`Auth failed: ${data.error_description}`);
    }
  }

  private async refreshToken(): Promise<void> {
    const res = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: CLIENT_ID,
          grant_type: "refresh_token",
          refresh_token: this.tokenData!.refresh_token,
          scope: SCOPES,
        }).toString(),
      },
    );
    const data = (await res.json()) as any;
    if (data.error) throw new Error(data.error_description);
    this.saveTokens({
      access_token: data.access_token,
      refresh_token: data.refresh_token || this.tokenData!.refresh_token,
      expires_at: Date.now() + data.expires_in * 1000,
    });
  }

  private async graphRequest(
    path: string,
    options: RequestInit = {},
  ): Promise<any> {
    if (this.tokenData && Date.now() > this.tokenData.expires_at - 60000)
      await this.refreshToken();
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

  private async graphRequestRaw(
    path: string,
    options: RequestInit = {},
  ): Promise<void> {
    if (this.tokenData && Date.now() > this.tokenData.expires_at - 60000)
      await this.refreshToken();
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
  }

  async getUnreadEmails(): Promise<Email[]> {
    console.log("📥 Fetching unread emails...");
    const data = await this.graphRequest(
      "/me/mailFolders/inbox/messages?$filter=isRead eq false&$top=50&$orderby=receivedDateTime asc&$select=id,subject,from,body,receivedDateTime",
    );
    const emails: Email[] = (data.value || []).map((msg: any) => ({
      id: msg.id,
      subject: msg.subject || "No Subject",
      sender: msg.from?.emailAddress?.name || "Unknown",
      senderEmail: msg.from?.emailAddress?.address || "",
      body: msg.body?.content
        ? msg.body.content
            .replace(/<[^>]*>/g, "")
            .replace(/\s+/g, " ")
            .trim()
            .slice(0, 2000)
        : "",
      receivedAt: msg.receivedDateTime || new Date().toISOString(),
    }));
    console.log(`✅ Found ${emails.length} unread emails`);
    return emails;
  }

  async replyToEmail(emailId: string, replyText: string): Promise<boolean> {
    try {
      await this.graphRequestRaw(`/me/messages/${emailId}/reply`, {
        method: "POST",
        body: JSON.stringify({
          message: { body: { contentType: "Text", content: replyText } },
        }),
      });
      return true;
    } catch (err) {
      console.error("Reply error:", err);
      return false;
    }
  }

  async openSharedMailbox(_sharedEmail: string): Promise<void> {
    console.log("✅ Using Graph API — no mailbox switching needed");
  }

  async close(): Promise<void> {}
}
