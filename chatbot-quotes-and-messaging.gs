/**
 * U.S. Stoneworks — Quote Chatbot (SMS / WhatsApp / Email) — No Twilio
 * ----------------------------------------------------------------------
 * v2: replaces Twilio with two lower-cost / self-hosted alternatives:
 *   - WhatsApp → Meta's own WhatsApp Cloud API, direct (no BSP markup)
 *   - SMS      → SMS Gateway for Android (self-hosted, open source —
 *                turns a spare Android phone + SIM into your SMS server)
 *   - Email    → unchanged: Gmail polling, same as before
 *
 * Pricing accuracy is unchanged: dollar amounts always come from the
 * spreadsheet, never from the AI. Claude only writes the intro sentence
 * and interprets loosely-worded messages.
 *
 * SETUP — see CHATBOT-SETUP.md for the full walkthrough. Script
 * Properties needed (Project Settings > Script Properties):
 *   ANTHROPIC_API_KEY   — from console.anthropic.com
 *   META_WHATSAPP_TOKEN — permanent access token from Meta Business Manager
 *   META_PHONE_NUMBER_ID — from your WhatsApp app's Meta dashboard
 *   META_VERIFY_TOKEN   — any string you make up yourself; enter the same
 *                         string in Meta's webhook config
 *   SMSGATE_USERNAME    — from the SMS Gateway Android app (Cloud Server)
 *   SMSGATE_PASSWORD    — from the SMS Gateway Android app (Cloud Server)
 */

const CLAUDE_MODEL = "claude-haiku-4-5-20251001"; // fast + inexpensive; swap for
                                                   // "claude-sonnet-5" for higher quality replies
const GMAIL_LABEL = "Chatbot-Replied";
const GMAIL_POLL_QUERY = "is:unread -label:" + GMAIL_LABEL;

// Quote MATCHING (findQuote_) never calls the Anthropic API, under any
// setting below — it's pure JS string/regex matching against the sheet.
// This toggle only controls the one-sentence friendly intro line that
// gets prepended to the price breakdown. Set to false to guarantee zero
// Anthropic API calls anywhere in this script (e.g. to run SMS/WhatsApp
// with no ANTHROPIC_API_KEY at all) — replies still contain the full,
// accurate quote breakdown either way.
const USE_AI_INTRO_SENTENCE = true;

// Paste your deployed Web App URL here (Deploy > Manage deployments) —
// used by registerSmsGatewayWebhook() to tell the SMS Gateway app where
// to send incoming-message webhooks.
const DEPLOYED_WEBAPP_URL = "";

// Shown in replies so people know email/text are both fine, and know
// where to go for anything the bot can't handle. Update the URL if you
// move to a custom domain later.
const CONTACT_URL = "https://us-stoneworks.github.io/us-stoneworks/#contact";
const CONTACT_PHONE = "(555) 010-2847"; // placeholder — update to the real number

function contactFooter_() {
  return `Email or text us anytime with questions — or visit ${CONTACT_URL} to reach us directly.`;
}

function prop_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  if (!v) throw new Error(`Script property ${name} is not set.`);
  return v;
}

// ============================================================
// Quote lookup — pure spreadsheet reads, no AI involved
// ============================================================

function getSheetData_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values
    .filter(row => row.some(cell => cell !== "" && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function findQuote_(messageText, senderPhone, senderEmail) {
  const quotes = getSheetData_("Quotes");
  const text = (messageText || "").toLowerCase();

  const idMatch = text.match(/\b(\d{2,6})\b/);
  if (idMatch) {
    const found = quotes.find(q => String(q["Quote ID"]) === idMatch[1]);
    if (found) return found;
  }

  if (senderPhone) {
    const digits = senderPhone.replace(/\D/g, "").slice(-10);
    const found = quotes.find(q =>
      String(q["Customer Phone"] || "").replace(/\D/g, "").slice(-10) === digits && digits.length === 10
    );
    if (found) return found;
  }

  if (senderEmail) {
    const found = quotes.find(q =>
      String(q["Customer Email"] || "").toLowerCase() === senderEmail.toLowerCase()
    );
    if (found) return found;
  }

  const found = quotes.find(q => {
    const addr = String(q["Address"] || "").toLowerCase();
    if (!addr) return false;
    const words = addr.split(/\s+/).filter(w => w.length > 2);
    return words.length > 0 && words.every(w => text.includes(w));
  });
  return found || null;
}

function getLineItems_(quoteId) {
  return getSheetData_("Line Items").filter(li => String(li["Quote ID"]) === String(quoteId));
}

function formatMoney_(n) {
  return "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function formatQuoteReply_(quote, lineItems) {
  const lines = [];
  lines.push(quote["Address"]);
  if (quote["City"]) lines.push(quote["City"]);
  lines.push("");

  lineItems.forEach(li => {
    const unitLabel = li["Unit"] === "LF" ? "lineal feet" : li["Unit"] === "EA" ? "" : li["Unit"];
    const qtyPart = unitLabel ? `${li["Qty"]} ${unitLabel} ` : "";
    lines.push(`${qtyPart}${li["Description"]} ${formatMoney_(li["Amount"])}`);
  });

  lines.push("");
  lines.push(`Total: ${formatMoney_(quote["Total"])}`);
  if (quote["Valid Until"]) lines.push(`(Quote valid until ${quote["Valid Until"]})`);
  lines.push("");
  lines.push(contactFooter_());
  return lines.join("\n");
}

// ============================================================
// Claude — used ONLY to interpret free-text intent / write the
// short conversational wrapper. Never asked to invent numbers.
// ============================================================

function askClaude_(systemPrompt, userMessage) {
  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": prop_("ANTHROPIC_API_KEY"),
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 300,
      system: systemPrompt,
      messages: [{ role: "user", content: userMessage }]
    }),
    muteHttpExceptions: true
  });

  const body = JSON.parse(response.getContentText());
  if (body.content && body.content[0] && body.content[0].text) {
    return body.content[0].text.trim();
  }
  return null;
}

function composeReplyWithQuote_(quote, lineItems, incomingMessage) {
  const breakdown = formatQuoteReply_(quote, lineItems);

  if (!USE_AI_INTRO_SENTENCE) return breakdown; // Anthropic API never called

  try {
    const intro = askClaude_(
      "You write ONE short, warm, professional sentence (under 25 words) introducing a price " +
      "quote breakdown that will be appended right after your sentence. Do not include any " +
      "dollar amounts or line items yourself — those are appended separately. No greeting like " +
      "'Hi', just the sentence. Reply with only that sentence, nothing else.",
      `Customer's message: "${incomingMessage}"\nProject address: ${quote["Address"]}, ${quote["City"]}`
    );
    return intro ? `${intro}\n\n${breakdown}` : breakdown;
  } catch (err) {
    return breakdown; // Claude unavailable — still send the accurate breakdown
  }
}

function noMatchReply_() {
  return "Thanks for reaching out! I couldn't find a quote matching that — could you send the " +
         "property address or your quote number? Email and text messages are both welcome, and " +
         `you can always reach us directly at ${CONTACT_PHONE} or ${CONTACT_URL}.`;
}

/**
 * Run this from the Apps Script editor (function dropdown > Run) to prove
 * quote matching and reply formatting work with zero Anthropic API calls —
 * it never touches askClaude_ or USE_AI_INTRO_SENTENCE. Check View > Logs.
 */
function testMatchingWithoutAI() {
  const samples = [
    "quote 388",
    "hey what's the price for 734 Rancho Cir",
    "checking on Fullerton job"
  ];
  samples.forEach(text => {
    const quote = findQuote_(text, null, null);
    Logger.log(`Message: "${text}"`);
    Logger.log(quote ? formatQuoteReply_(quote, getLineItems_(quote["Quote ID"])) : "NO MATCH");
    Logger.log("---");
  });
}

// ============================================================
// WhatsApp — Meta Cloud API, direct (no Twilio/BSP)
// ============================================================

/** Meta calls this with GET once, to verify you own the webhook URL. */
function doGet(e) {
  const mode = e.parameter["hub.mode"];
  const token = e.parameter["hub.verify_token"];
  const challenge = e.parameter["hub.challenge"];

  if (mode === "subscribe" && token === PropertiesService.getScriptProperties().getProperty("META_VERIFY_TOKEN")) {
    return ContentService.createTextOutput(challenge);
  }
  return ContentService.createTextOutput("U.S. Stoneworks chatbot endpoint is live.");
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput("OK");
  }

  if (body.object === "whatsapp_business_account") {
    handleWhatsAppWebhook_(body);
  } else if (body.event === "sms:received") {
    handleSmsWebhook_(body);
  }

  // Both Meta and the SMS Gateway app just need a 200 OK acknowledgement.
  return ContentService.createTextOutput("OK");
}

function handleWhatsAppWebhook_(body) {
  (body.entry || []).forEach(entry => {
    (entry.changes || []).forEach(change => {
      const messages = change.value && change.value.messages;
      if (!messages) return; // status callbacks (delivered/read) have no "messages" field

      messages.forEach(msg => {
        if (msg.type !== "text") return;
        const fromPhone = msg.from; // digits only, no "+"
        const text = msg.text.body;

        const quote = findQuote_(text, fromPhone, null);
        const reply = quote
          ? composeReplyWithQuote_(quote, getLineItems_(quote["Quote ID"]), text)
          : noMatchReply_();

        sendWhatsAppMessage_(fromPhone, reply);
      });
    });
  });
}

function sendWhatsAppMessage_(toPhone, text) {
  const phoneId = prop_("META_PHONE_NUMBER_ID");
  UrlFetchApp.fetch(`https://graph.facebook.com/v23.0/${phoneId}/messages`, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + prop_("META_WHATSAPP_TOKEN") },
    payload: JSON.stringify({
      messaging_product: "whatsapp",
      to: toPhone,
      type: "text",
      text: { body: text }
    }),
    muteHttpExceptions: true
  });
}

// ============================================================
// SMS — SMS Gateway for Android (self-hosted, open source)
// https://sms-gate.app — turns a spare Android phone into the sender
// ============================================================

function handleSmsWebhook_(body) {
  const payload = body.payload || {};
  const messageText = payload.message;
  const fromPhone = payload.phoneNumber;
  if (!messageText || !fromPhone) return;

  const quote = findQuote_(messageText, fromPhone, null);
  const reply = quote
    ? composeReplyWithQuote_(quote, getLineItems_(quote["Quote ID"]), messageText)
    : noMatchReply_();

  sendSms_(fromPhone, reply);
}

function sendSms_(toPhone, text) {
  const auth = Utilities.base64Encode(prop_("SMSGATE_USERNAME") + ":" + prop_("SMSGATE_PASSWORD"));
  UrlFetchApp.fetch("https://api.sms-gate.app/3rdparty/v1/message", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Basic " + auth },
    payload: JSON.stringify({
      textMessage: { text: text },
      phoneNumbers: [toPhone]
    }),
    muteHttpExceptions: true
  });
}

/**
 * Run this ONCE (from the Apps Script editor's function dropdown) after
 * you've set DEPLOYED_WEBAPP_URL and the SMSGATE_* script properties.
 * Registers your Web App URL with the SMS Gateway cloud service so it
 * knows where to forward incoming SMS.
 */
function registerSmsGatewayWebhook() {
  if (!DEPLOYED_WEBAPP_URL) throw new Error("Set DEPLOYED_WEBAPP_URL at the top of this file first.");
  const auth = Utilities.base64Encode(prop_("SMSGATE_USERNAME") + ":" + prop_("SMSGATE_PASSWORD"));
  const res = UrlFetchApp.fetch("https://api.sms-gate.app/3rdparty/v1/webhooks", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Basic " + auth },
    payload: JSON.stringify({
      id: "us-stoneworks-chatbot",
      url: DEPLOYED_WEBAPP_URL,
      event: "sms:received"
    }),
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}

// ============================================================
// Email — polls Gmail on a time-driven trigger (unchanged)
// ============================================================

function checkInboxAndReply() {
  ensureGmailLabel_();
  const threads = GmailApp.search(GMAIL_POLL_QUERY, 0, 20);

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const last = messages[messages.length - 1];
    if (last.isDraft()) return;

    const fromEmail = extractEmailAddress_(last.getFrom());
    const body = last.getPlainBody();

    const quote = findQuote_(body + " " + last.getSubject(), null, fromEmail);
    const replyText = quote
      ? composeReplyWithQuote_(quote, getLineItems_(quote["Quote ID"]), body)
      : noMatchReply_();

    last.reply(replyText);
    thread.addLabel(GmailApp.getUserLabelByName(GMAIL_LABEL));
    last.markRead();
  });
}

function extractEmailAddress_(fromHeader) {
  const match = fromHeader.match(/<(.+)>/);
  return match ? match[1] : fromHeader;
}

function ensureGmailLabel_() {
  if (!GmailApp.getUserLabelByName(GMAIL_LABEL)) {
    GmailApp.createLabel(GMAIL_LABEL);
  }
}

/** Run this once from the Apps Script editor to schedule inbox checks. */
function enableEmailTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "checkInboxAndReply") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkInboxAndReply").timeBased().everyMinutes(5).create();
}
