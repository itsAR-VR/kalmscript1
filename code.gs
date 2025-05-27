/**
 * The exact subject for Outreach and for thread-matching in all helpers.
 */
const OUTREACH_SUBJECT = `Hey We'd love to send you some product! // kalm wellness`;

// Number of days to wait before each follow-up email is sent
const FIRST_FU_DELAY_DAYS  = 2;
const SECOND_FU_DELAY_DAYS = 4;
const THIRD_FU_DELAY_DAYS  = 7;
const FOURTH_FU_DELAY_DAYS = 12;


/**
 * Installable onEdit trigger: fires on ANY sheet when "Status" is edited.
 * Now only dispatches on the tag(s) you just added.
 */
function onEditTrigger(e) {
  if (!e || !e.range) return;
  const sh   = e.range.getSheet();
  const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  // 1) Find Status column
  const statusCol = hdrs.indexOf('Status') + 1;
  if (statusCol < 1 || e.range.getColumn() !== statusCol) return;

  // 2) Grab the new & old values
  const newStatus = e.value    || '';
  const oldStatus = e.oldValue || '';

  // 3) Compute which tags were just added
  const newTags = newStatus.split(',').map(t => t.trim()).filter(Boolean);
  const oldTags = oldStatus.split(',').map(t => t.trim()).filter(Boolean);
  const additions = newTags.filter(t => !oldTags.includes(t));
  if (!additions.length) return;
  Logger.log('Tags added: %s', additions.join(', '));

  // 4) Find Name & Email columns
  const nameCol  = hdrs.indexOf('First/Last Name') + 1;
  const emailCol = hdrs.indexOf('Email')           + 1;
  if (nameCol < 1 || emailCol < 1) {
    throw new Error('Headers required: First/Last Name, Email, Status');
  }

  // 5) Read that row’s data
  const row    = e.range.getRow();
  const vals   = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const full   = vals[nameCol - 1] || '';
  const first  = full.split(/\s+/)[0];
  const email  = vals[emailCol - 1];
  if (!email) return;

  // 6) Dispatch exactly for each newly added tag
  additions.forEach(tag => {
    switch (tag) {
      case 'Outreach':
        Logger.log('Dispatching Outreach for %s', email);
        sendInitialForRow(email, first);
        break;
      case '1st Follow Up':
        Logger.log('Dispatching 1st Follow-Up for %s', email);
        sendFirstFollowUpForRow(email, first);
        break;
      case '2nd Follow Up':
        Logger.log('Dispatching 2nd Follow-Up for %s', email);
        sendSecondFollowUpForRow(email, first);
        break;
      case '3rd Follow Up':
        Logger.log('Dispatching 3rd Follow-Up for %s', email);
        sendThirdFollowUpForRow(email, first);
        break;
      case '4th Follow Up':
        Logger.log('Dispatching 4th Follow-Up for %s', email);
        sendFourthFollowUpForRow(email, first);
        break;

      default:
        Logger.log('Unknown tag "%s"; skipping', tag);
    }
  });
}


/**
 * 1) Outreach: one-off sendEmail using OutreachTemplate.html
 */
function sendInitialForRow(email, firstName) {
  const subject  = OUTREACH_SUBJECT;
  const tpl      = HtmlService.createTemplateFromFile('OutreachTemplate');
  tpl.firstName  = firstName;
  const htmlBody = tpl.evaluate().getContent();
  const textBody = `Hi ${firstName},\n\nThanks for connecting—here’s the info we discussed.\n\nBest,\nKam Ordonez`;

  MailApp.sendEmail({
    to:       email,
    subject:  subject,
    body:     textBody,
    htmlBody: htmlBody
  });
  Logger.log('Outreach sent to %s with subject "%s"', email, subject);
}


/**
 * 2) First Follow-Up: advanced threaded reply that sets To: explicitly.
 */
function sendFirstFollowUpForRow(email, firstName) {
  Logger.log('▶ Enter sendFirstFollowUpForRow; email=%s, firstName=%s', email, firstName);

  const subject = OUTREACH_SUBJECT;
  const query   = `in:anywhere to:${email} subject:"${subject}"`;
  Logger.log('Search query: %s', query);
  
  const threads = GmailApp.search(query);
  Logger.log('Search returned %d thread(s)', threads.length);
  if (!threads.length) return;

  const thread   = threads[0];
  const lastMsg  = thread.getMessages().pop();
  const rawOrig  = lastMsg.getRawContent();
  const inReplyTo= (rawOrig.match(/^Message-ID:\s*(<[^>]+>)/mi) || [])[1];
  if (!inReplyTo) {
    Logger.log('❌ No Message-ID header found; aborting.');
    return;
  }

  // Render HTML & text
  const tpl      = HtmlService.createTemplateFromFile('FirstFollowUpTemplate');
  tpl.firstName  = firstName;
  const htmlBody = tpl.evaluate().getContent();
  const textBody = `Hi ${firstName},\n\nJust checking in—any questions?\n\nBest,\nKam Ordonez`;

  // Build raw RFC-2822 reply
  const raw = buildRawMessage_(email, `Re: ${subject}`, textBody, htmlBody, inReplyTo);
  Gmail.Users.Messages.send({ threadId: thread.getId(), raw: raw }, 'me');
  Logger.log('✅ 1st FU sent via Advanced API to %s in thread %s', email, thread.getId());
}

/**
 * 3) Second Follow-Up: same advanced send logic.
 */
function sendSecondFollowUpForRow(email, firstName) {
  Logger.log('▶ Enter sendSecondFollowUpForRow; email=%s, firstName=%s', email, firstName);

  const subject = OUTREACH_SUBJECT;
  const query   = `in:anywhere to:${email} subject:"${subject}"`;
  Logger.log('Search query: %s', query);
  
  const threads = GmailApp.search(query);
  Logger.log('Search returned %d thread(s)', threads.length);
  if (!threads.length) return;

  const thread   = threads[0];
  const lastMsg  = thread.getMessages().pop();
  const rawOrig  = lastMsg.getRawContent();
  const inReplyTo= (rawOrig.match(/^Message-ID:\s*(<[^>]+>)/mi) || [])[1];
  if (!inReplyTo) {
    Logger.log('❌ No Message-ID header found; aborting.');
    return;
  }

  // Render HTML & text
  const tpl      = HtmlService.createTemplateFromFile('SecondFollowUpTemplate');
  tpl.firstName  = firstName;
  const htmlBody = tpl.evaluate().getContent();
  const textBody = `Hi ${firstName},\n\nJust checking-in let me know if you need anything else!\n\nBest,\nKam Ordonez`;

  // Build and send raw reply
  const raw = buildRawMessage_(email, `Re: ${subject}`, textBody, htmlBody, inReplyTo);
  Gmail.Users.Messages.send({ threadId: thread.getId(), raw: raw }, 'me');
  Logger.log('✅ 2nd FU sent via Advanced API to %s in thread %s', email, thread.getId());
}

function sendThirdFollowUpForRow(email, firstName) {
  Logger.log('▶ Enter sendThirdFollowUpForRow; email=%s, firstName=%s', email, firstName);

  const subject = OUTREACH_SUBJECT;
  const query   = `in:anywhere to:${email} subject:"${subject}"`;
  Logger.log('Search query: %s', query);

  const threads = GmailApp.search(query);
  Logger.log('Search returned %d thread(s)', threads.length);
  if (!threads.length) return;

  const thread   = threads[0];
  const lastMsg  = thread.getMessages().pop();
  const rawOrig  = lastMsg.getRawContent();
  const inReplyTo= (rawOrig.match(/^Message-ID:\s*(<[^>]+>)/mi) || [])[1];
  if (!inReplyTo) {
    Logger.log('❌ No Message-ID header found; aborting third follow-up.');
    return;
  }

  const tpl      = HtmlService.createTemplateFromFile('ThirdFollowUpTemplate');
  tpl.firstName  = firstName;
  const htmlBody = tpl.evaluate().getContent();
  const textBody = `Hi ${firstName},\n\nQuick nudge—your complimentary Kalm mouth‑tape pack is still reserved for you. Just reply with your address and I’ll ship it right away!\n\nWarmly,\nKam Ordonez`;

  const raw = buildRawMessage_(email, `Re: ${subject}`, textBody, htmlBody, inReplyTo);
  Gmail.Users.Messages.send({ threadId: thread.getId(), raw: raw }, 'me');
  Logger.log('✅ 3rd FU sent via Advanced API to %s in thread %s', email, thread.getId());
}

/**
 * 5) Fourth (Final) Follow‑Up: graceful close‑out 10–12 days later.
 */
function sendFourthFollowUpForRow(email, firstName) {
  Logger.log('▶ Enter sendFourthFollowUpForRow; email=%s, firstName=%s', email, firstName);

  const subject = OUTREACH_SUBJECT;
  const query   = `in:anywhere to:${email} subject:"${subject}"`;
  Logger.log('Search query: %s', query);

  const threads = GmailApp.search(query);
  Logger.log('Search returned %d thread(s)', threads.length);
  if (!threads.length) return;

  const thread   = threads[0];
  const lastMsg  = thread.getMessages().pop();
  const rawOrig  = lastMsg.getRawContent();
  const inReplyTo= (rawOrig.match(/^Message-ID:\s*(<[^>]+>)/mi) || [])[1];
  if (!inReplyTo) {
    Logger.log('❌ No Message-ID header found; aborting fourth follow-up.');
    return;
  }

  const tpl      = HtmlService.createTemplateFromFile('FourthFollowUpTemplate');
  tpl.firstName  = firstName;
  const htmlBody = tpl.evaluate().getContent();
  const textBody = `Hi ${firstName},\n\nThis is my last check‑in for now. If calmer, clearer sleep isn’t on your radar yet, no worries—just reply “later”. Otherwise, send your address anytime and I’ll pop your free sample in the mail.\n\nFind your Kalm,\nKam Ordonez`;

  const raw = buildRawMessage_(email, `Re: ${subject}`, textBody, htmlBody, inReplyTo);
  Gmail.Users.Messages.send({ threadId: thread.getId(), raw: raw }, 'me');
  Logger.log('✅ 4th FU sent via Advanced API to %s in thread %s', email, thread.getId());
}

/**
 * Helper: builds a base64-URL-encoded RFC-2822 multipart/alternative reply
 * with To:, Subject:, In-Reply-To, and References headers.
 */
function buildRawMessage_(to, subject, textBody, htmlBody, inReplyTo) {
  const nl       = '\r\n';
  const boundary = '----=_Boundary_' + Date.now();
  let msg =
    `To: ${to}` + nl +
    `Subject: ${subject}` + nl +
    `In-Reply-To: ${inReplyTo}` + nl +
    `References: ${inReplyTo}` + nl +
    `MIME-Version: 1.0` + nl +
    `Content-Type: multipart/alternative; boundary="${boundary}"` + nl + nl +
    `--${boundary}` + nl +
    `Content-Type: text/plain; charset="UTF-8"` + nl + nl +
    textBody + nl + nl +
    `--${boundary}` + nl +
    `Content-Type: text/html; charset="UTF-8"` + nl + nl +
    htmlBody + nl + nl +
    `--${boundary}--`;
  return Utilities.base64EncodeWebSafe(msg);
}

/**
 * Automatically send follow-up emails if contacts haven't replied.
 * Intended to run daily via a time-based Apps Script trigger.
 */
function autoSendFollowUps() {
  const sh   = SpreadsheetApp.getActiveSheet();
  const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const nameCol   = hdrs.indexOf('First/Last Name') + 1;
  const emailCol  = hdrs.indexOf('Email') + 1;
  const statusCol = hdrs.indexOf('Status') + 1;
  if (nameCol < 1 || emailCol < 1 || statusCol < 1) {
    throw new Error('Headers required: First/Last Name, Email, Status');
  }

  const numRows = sh.getLastRow() - 1;
  if (numRows < 1) return;
  const data = sh.getRange(2, 1, numRows, sh.getLastColumn()).getValues();

  data.forEach((vals, idx) => {
    const row    = idx + 2;
    const email  = vals[emailCol - 1];
    if (!email) return;
    const full   = vals[nameCol - 1] || '';
    const first  = full.split(/\s+/)[0];
    let status   = vals[statusCol - 1] || '';
    const tags   = status.split(',').map(t => t.trim()).filter(Boolean);

    const query   = `in:anywhere to:${email} subject:"${OUTREACH_SUBJECT}"`;
    const threads = GmailApp.search(query);
    if (!threads.length) return;

    const thread   = threads[0];
    const lastMsg  = thread.getMessages().pop();
    const fromAddr = lastMsg.getFrom();
    if (fromAddr && fromAddr.toLowerCase().includes(email.toLowerCase())) return;

    const daysSince = (Date.now() - lastMsg.getDate().getTime()) / 86400000;

    if (!tags.includes('1st Follow Up Sent') && daysSince >= FIRST_FU_DELAY_DAYS) {
      sendFirstFollowUpForRow(email, first);
      tags.push('1st Follow Up Sent');
    } else if (
      tags.includes('1st Follow Up Sent') &&
      !tags.includes('2nd Follow Up Sent') &&
      daysSince >= SECOND_FU_DELAY_DAYS
    ) {
      sendSecondFollowUpForRow(email, first);
      tags.push('2nd Follow Up Sent');
    } else if (
      tags.includes('2nd Follow Up Sent') &&
      !tags.includes('3rd Follow Up Sent') &&
      daysSince >= THIRD_FU_DELAY_DAYS
    ) {
      sendThirdFollowUpForRow(email, first);
      tags.push('3rd Follow Up Sent');
    } else if (
      tags.includes('3rd Follow Up Sent') &&
      !tags.includes('4th Follow Up Sent') &&
      daysSince >= FOURTH_FU_DELAY_DAYS
    ) {
      sendFourthFollowUpForRow(email, first);
      tags.push('4th Follow Up Sent');
    }

    const newStatus = tags.join(', ');
    if (newStatus !== status) {
      sh.getRange(row, statusCol).setValue(newStatus);
    }
  });
}
