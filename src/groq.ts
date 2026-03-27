import Groq from "groq-sdk";
import { Email } from "./outlook";

const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

export type ReplyType = "approved" | "denied" | "more_info" | "acknowledgement";

export interface EmailProcessResult {
  replyType: ReplyType;
  draft: string;
}

const CLASSIFY_AND_DRAFT_PROMPT = `You are a leave services officer at NRS. Classify and draft a brief, formal reply to a staff leave request.

Reply rules: address sender by name, 3-5 sentences, end with:
Regards,
Gideon Buba,
Employee Service Delivery

Return ONLY this JSON (no markdown):
{"type":"acknowledgement","reply":"Your reply here"}

Types: approved (clear request), denied (excessive/bad timing/incomplete), more_info (missing dates/type/duration), acknowledgement (needs manual review)`;

const FALLBACK_REPLY =
  "Thank you for your leave request. It has been received and is currently being reviewed. We will get back to you shortly.\n\nRegards,\nGideon Buba,\nEmployee Service Delivery";

export function shouldSkipEmail(email: Email): boolean {
  const subject = email.subject.toLowerCase();
  return (
    subject.startsWith("re:") ||
    subject.startsWith("fw:") ||
    subject.startsWith("fwd:")
  );
}

export async function processEmail(email: Email): Promise<EmailProcessResult> {
  const truncatedBody = email.body.slice(0, 800);

  const prompt = `
Sender: ${email.sender} (${email.senderEmail})
Subject: ${email.subject}
Email body:
${truncatedBody}

Classify this leave request and draft an appropriate reply.
`.trim();

  const completion = await groq.chat.completions.create({
    model: "llama-3.1-8b-instant",
    messages: [
      { role: "system", content: CLASSIFY_AND_DRAFT_PROMPT },
      { role: "user", content: prompt },
    ],
    temperature: 0.3,
    max_tokens: 400,
  });

  const raw = completion.choices[0]?.message?.content?.trim() || "";

  try {
    const clean = raw.replace(/```json|```/g, "").trim();
    const parsed = JSON.parse(clean);

    const validTypes: ReplyType[] = [
      "approved",
      "denied",
      "more_info",
      "acknowledgement",
    ];
    const replyType: ReplyType = validTypes.includes(parsed.type)
      ? parsed.type
      : "acknowledgement";
    const draft: string = parsed.reply?.trim() || FALLBACK_REPLY;

    return { replyType, draft };
  } catch {
    console.error("Failed to parse LLM response:", raw);
    return { replyType: "acknowledgement", draft: FALLBACK_REPLY };
  }
}

const SIGN_OFF = "\n\nRegards,\nGideon Buba,\nEmployee Service Delivery";

export function draftReply(email: Email, replyType: ReplyType): string {
  const name = email.sender?.split(" ")[0] || "Sir/Ma";
  const templates: Record<ReplyType, string> = {
    approved: `Dear ${name},\n\nThank you for your leave request. After review, your request has been approved. Please ensure a proper handover before your leave commences.${SIGN_OFF}`,
    denied: `Dear ${name},\n\nThank you for your leave request. Unfortunately, we are unable to approve your request at this time. Please feel free to reapply at a more suitable period.${SIGN_OFF}`,
    more_info: `Dear ${name},\n\nThank you for reaching out. To process your leave request, we require additional details such as the leave type, start date, end date, and duration. Kindly provide these at your earliest convenience.${SIGN_OFF}`,
    acknowledgement: `Dear ${name},\n\nThank you for your leave request. It has been received and is currently under review. We will revert to you as soon as possible.${SIGN_OFF}`,
  };
  return templates[replyType] || FALLBACK_REPLY;
}

export async function classifyEmail(email: Email): Promise<ReplyType> {
  const { replyType } = await processEmail(email);
  return replyType;
}
