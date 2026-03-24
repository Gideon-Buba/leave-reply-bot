import Groq from "groq-sdk";
import { Email } from "./outlook";

const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

const SYSTEM_PROMPT = `You are a professional leave services officer at NRS (Nigerian government organization). 
You draft concise, formal, and polite email replies to staff leave requests.

Your replies should:
- Be professional but warm
- Address the sender by name if available
- Be brief (3-5 sentences max)
- Reference the specific leave type/dates mentioned if clear
- End with a standard sign-off: "Regards,\nGideon Buba,\nOII, Leave Unit"

You will be given the email content and the type of response needed.
Return ONLY the email reply body — no subject line, no extra commentary.`;

export type ReplyType = "approved" | "denied" | "more_info" | "acknowledgement";

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
${email.body}

Task: ${replyInstructions[replyType]}
  `.trim();

  const completion = await groq.chat.completions.create({
    model: "llama-3.3-70b-versatile",
    messages: [
      { role: "system", content: SYSTEM_PROMPT },
      { role: "user", content: prompt },
    ],
    temperature: 0.4,
    max_tokens: 300,
  });

  return (
    completion.choices[0]?.message?.content?.trim() ||
    "Thank you for your leave request. We will get back to you shortly.\n\nRegards,\nLeave Services Unit"
  );
}

export async function classifyEmail(email: Email): Promise<ReplyType> {
  const prompt = `
Classify this email into one of these reply types:
- "approved": Staff is requesting leave and it seems straightforward to approve
- "denied": Request seems problematic (too long, bad timing mentioned, incomplete forms)
- "more_info": Request is missing key details (dates, leave type, duration)
- "acknowledgement": Complex request that needs more review

Email subject: ${email.subject}
Email body: ${email.body.slice(0, 500)}

Reply with ONLY one of these exact words: approved, denied, more_info, acknowledgement
  `.trim();

  const completion = await groq.chat.completions.create({
    model: "llama-3.3-70b-versatile",
    messages: [{ role: "user", content: prompt }],
    temperature: 0,
    max_tokens: 10,
  });

  const result = completion.choices[0]?.message?.content?.trim().toLowerCase();
  const validTypes: ReplyType[] = [
    "approved",
    "denied",
    "more_info",
    "acknowledgement",
  ];
  return validTypes.includes(result as ReplyType)
    ? (result as ReplyType)
    : "acknowledgement";
}
