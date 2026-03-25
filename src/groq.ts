import Groq from "groq-sdk";
import { Email } from "./outlook";

const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

export type ReplyType = "approved" | "denied" | "more_info" | "acknowledgement";

export interface EmailProcessResult {
  replyType: ReplyType;
  draft: string;
}

const CLASSIFY_AND_DRAFT_PROMPT = `You are a professional leave services officer at NRS (Nigerian government organization).
You classify and draft concise, formal, and polite email replies to staff leave requests.

Your replies should:
- Be professional but warm
- Address the sender by name if available
- Be brief (3-5 sentences max)
- Reference the specific leave type/dates mentioned if clear
- End with a standard sign-off: "Regards,\nGideon Buba,\nOII, Leave Unit"

Return ONLY a JSON object in this exact format, no extra text, no markdown:
{"type":"acknowledgement","reply":"Your reply here"}

Where "type" is one of: approved, denied, more_info, acknowledgement
- approved: straightforward leave request, safe to approve
- denied: problematic request (excessive duration, bad timing, incomplete forms)
- more_info: missing key details (dates, leave type, duration)
- acknowledgement: complex or unclear request that needs manual review`;

const DRAFT_ONLY_PROMPT = `You are a professional leave services officer at NRS (Nigerian government organization).
You draft concise, formal, and polite email replies to staff leave requests.

Your replies should:
- Be professional but warm
- Address the sender by name if available
- Be brief (3-5 sentences max)
- Reference the specific leave type/dates mentioned if clear
- End with a standard sign-off: "Regards,\nGideon Buba,\nOII, Leave Unit"

Return ONLY the plain email reply body — no subject line, no JSON, no markdown, no extra commentary.`;

const FALLBACK_REPLY =
  "Thank you for your leave request. It has been received and is currently being reviewed. We will get back to you shortly.\n\nRegards,\nGideon Buba,\nOII, Leave Unit";

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
    model: "llama-3.3-70b-versatile",
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

export async function draftReply(
  email: Email,
  replyType: ReplyType,
): Promise<string> {
  const replyInstructions: Record<ReplyType, string> = {
    approved:
      "Draft an approval reply confirming their leave request has been approved.",
    denied:
      "Draft a polite denial reply explaining their leave request cannot be approved at this time.",
    more_info:
      "Draft a reply requesting more information or clarification about their leave request.",
    acknowledgement:
      "Draft an acknowledgement reply confirming their request has been received and is being reviewed.",
  };

  const prompt = `
Sender: ${email.sender} (${email.senderEmail})
Subject: ${email.subject}
Email body:
${email.body.slice(0, 800)}

Task: ${replyInstructions[replyType]}
`.trim();

  const completion = await groq.chat.completions.create({
    model: "llama-3.3-70b-versatile",
    messages: [
      { role: "system", content: DRAFT_ONLY_PROMPT },
      { role: "user", content: prompt },
    ],
    temperature: 0.3,
    max_tokens: 400,
  });

  return completion.choices[0]?.message?.content?.trim() || FALLBACK_REPLY;
}

export async function classifyEmail(email: Email): Promise<ReplyType> {
  const { replyType } = await processEmail(email);
  return replyType;
}
