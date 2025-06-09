/**
 * The exact subject for Outreach and for thread-matching in all helpers.
 */
const OUTREACH_SUBJECT = `Hey We'd love to send you some product! // kalm wellness`;

// Name of the sheet containing outreach contacts
const TARGET_SHEET_NAME = 'Influencer PR';

// Email address used when sending messages. Set this to the Gmail
// account that owns the script. Using one address avoids issues with
// aliases and makes reply detection reliable.
const FROM_ADDRESS = 'creators@clubkalm.com';

// Background color used when marking any reply in the sheet.
const NEW_RESPONSE_COLOR = 'red';

// Background color used when a contact is moved to DM.
const MOVED_TO_DM_COLOR = '#d9d9d9';

// Script property key used to control automatic sending of follow-ups.
const AUTO_SEND_ENABLED_PROP = 'AutoSendEnabled';

// Prefix used to build links back to Gmail threads in the Reply Status column.
const GMAIL_THREAD_LINK_PREFIX = 'https://mail.google.com/mail/u/0/#inbox/';

// Number of minutes to wait before each follow-up email is sent.
// These were previously day-based delays.  For production, keep the
// minute values equivalent to the desired day delays (e.g. 2 days =
// 2 * 24 * 60 = 2880).  During testing you can shorten these values
// to just a few minutes for faster feedback.
const FIRST_FU_DELAY_MINUTES  = 2  * 24 * 60;  // 2 days
const SECOND_FU_DELAY_MINUTES = 4  * 24 * 60;  // 4 days
const THIRD_FU_DELAY_MINUTES  = 7  * 24 * 60;  // 7 days
const FOURTH_FU_DELAY_MINUTES = 12 * 24 * 60;  // 12 days

/**
 * Enable or disable automatic follow-up sending.
 *
 * @param {boolean} enabled TRUE to enable auto-send, FALSE to disable.
 */
function setAutoSendEnabled(enabled) {
  PropertiesService.getScriptProperties()
    .setProperty(AUTO_SEND_ENABLED_PROP, enabled ? 'true' : 'false');
}

/**
 * Check if automatic sending of follow-ups is enabled.
 *
 * @return {boolean} TRUE if enabled, otherwise FALSE.
 */
function isAutoSendEnabled() {
  const val = PropertiesService.getScriptProperties()
    .getProperty(AUTO_SEND_ENABLED_PROP);
  return val === 'true';
}

/**
 * Toggle automatic sending of follow-ups on or off.
 * Can be run from Extensions → Macros or a sheet button.
 */
function toggleAutoSendEnabled() {
  const enabled = !isAutoSendEnabled();
  setAutoSendEnabled(enabled);
  SpreadsheetApp.getUi().alert(
    `Automatic follow-ups ${enabled ? 'enabled' : 'disabled'}.`,
  );
}

/**
 * Installable onEdit trigger: fires on ANY sheet when "Status" is edited.
 * Now only dispatches on the tag(s) you just added.
 */
function onEditTrigger(e) {
  if (!e || !e.range) return;
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sh || e.range.getSheet().getName() !== TARGET_SHEET_NAME) return;
  const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const stageCol = hdrs.indexOf('Stage') + 1;

  // 1) Find Status column
  const statusCol = hdrs.indexOf('Status') + 1;
  if (statusCol < 1 || e.range.getColumn() !== statusCol) return;

  // 2) Grab the new & old values
  const newStatus = e.value    || '';
  const oldStatus = e.oldValue || '';

  // 3) Compute which tags were just added
  const newTags = newStatus.split(',').map(t => t.trim()).filter(Boolean);
  const oldTags = oldStatus.split(',').map(t => t.trim()).filter(Boolean);
  if (!newTags.includes('Outreach')) {
    Logger.log('Row no longer tagged Outreach; skipping trigger.');
    return;
  }
  const additions = newTags.filter(t => !oldTags.includes(t));
  if (!additions.length) return;
  Logger.log('Tags added: %s', additions.join(', '));

  // 4) Find Name & Email columns
  const firstNameCol = hdrs.indexOf('First Name') + 1;
  const lastNameCol  = hdrs.indexOf('Last Name') + 1;
  const emailCol     = hdrs.indexOf('Email') + 1;
  const threadIdCol  = hdrs.indexOf('Thread ID') + 1;
  if (firstNameCol < 1 || lastNameCol < 1 || emailCol < 1 || threadIdCol < 1) {
    throw new Error('Headers required: First Name, Last Name, Email, Status, Thread ID');
  }

  // 5) Read that row’s data
  const row    = e.range.getRow();
  const vals   = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const first  = (vals[firstNameCol - 1] || '').toString();
  const last   = (vals[lastNameCol - 1]  || '').toString();
  const email  = vals[emailCol - 1];
  const threadId = vals[threadIdCol - 1];
  if (!email) return;
  if (!first && !last) return;

  // 6) Dispatch exactly for each newly added tag
  additions.forEach(tag => {
    switch (tag) {
      case 'Outreach':
        Logger.log('Dispatching Outreach for %s', email);
        sendInitialForRow(email, first, row);
        if (stageCol > 0) {
          sh.getRange(row, stageCol).setValue('Outreach');
        }
        break;
      case '1st Follow Up':
        Logger.log('Dispatching 1st Follow-Up for %s', email);
        sendFirstFollowUpForRow(email, first, threadId);
        if (stageCol > 0) {
          sh.getRange(row, stageCol).setValue('Follow Up 1');
        }
        break;
      case '2nd Follow Up':
        Logger.log('Dispatching 2nd Follow-Up for %s', email);
        sendSecondFollowUpForRow(email, first, threadId);
        if (stageCol > 0) {
          sh.getRange(row, stageCol).setValue('Follow Up 2');
        }
        break;
      case '3rd Follow Up':
        Logger.log('Dispatching 3rd Follow-Up for %s', email);
        sendThirdFollowUpForRow(email, first, threadId);
        if (stageCol > 0) {
          sh.getRange(row, stageCol).setValue('Follow Up 3');
        }
        break;
      case '4th Follow Up':
        Logger.log('Dispatching 4th Follow-Up for %s', email);
        sendFourthFollowUpForRow(email, first, threadId);
        if (stageCol > 0) {
          sh.getRange(row, stageCol).setValue('Follow Up 4');
        }
        break;

      default:
        Logger.log('Unknown tag "%s"; skipping', tag);
    }
  });
}


/**
 * 1) Outreach: one-off sendEmail using OutreachTemplate.html
 */
function sendInitialForRow(email, firstName, rowIndex) {
  const subject  = OUTREACH_SUBJECT;
  const tpl      = HtmlService.createTemplateFromFile('OutreachTemplate');
  tpl.firstName  = firstName;
  const htmlBody = tpl.evaluate().getContent();
  const textBody = `Hi ${firstName},\n\nThanks for connecting—here’s the info we discussed.\n\nBest,\nKam Ordonez`;

  const raw = buildRawMessage_(email, subject, textBody, htmlBody);
  const response = Gmail.Users.Messages.send({ raw: raw }, 'me');

  if (rowIndex) {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const sh   = ss.getSheetByName(TARGET_SHEET_NAME);
    if (sh) {
      const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const threadCol = hdrs.indexOf('Thread ID') + 1;
      if (threadCol > 0) {
        sh.getRange(rowIndex, threadCol).setValue(response.threadId);
      }
    }
  }

  Logger.log('Outreach sent via Advanced API to %s with subject "%s"', email, subject);
}

/**
 * Send an Outreach email for the row the user currently has selected.
 * Verifies the selection is on the target sheet and extracts the name
 * and email columns before delegating to {@link sendInitialForRow}.
 * Optionally tags the Status column with "Outreach" so the sheet
 * reflects that the initial message was sent.
 */
function startOutreachForSelectedRow() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh || sh.getName() !== TARGET_SHEET_NAME) {
    SpreadsheetApp.getUi()
      .alert(`Please run this from the "${TARGET_SHEET_NAME}" sheet.`);
    return;
  }

  const range = sh.getActiveRange();
  if (!range) return;
  const row = range.getRow();
  if (row <= 1) return;

  const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const firstNameCol = hdrs.indexOf('First Name') + 1;
  const lastNameCol  = hdrs.indexOf('Last Name') + 1;
  const emailCol  = hdrs.indexOf('Email') + 1;
  const statusCol = hdrs.indexOf('Status') + 1;
  const stageCol  = hdrs.indexOf('Stage') + 1;
  if (firstNameCol < 1 || lastNameCol < 1 || emailCol < 1) {
    SpreadsheetApp.getUi().alert('Headers required: First Name, Last Name and Email');
    return;
  }

  const vals   = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const first  = (vals[firstNameCol - 1] || '').toString();
  const last   = (vals[lastNameCol - 1]  || '').toString();
  const email  = vals[emailCol - 1];
  if (!email) {
    SpreadsheetApp.getUi().alert('No email found for the selected row.');
    return;
  }

  if (!first && !last) {
    SpreadsheetApp.getUi().alert(
      'First Name and Last Name are blank for the selected row.',
    );
    return;
  }

  sendInitialForRow(email, first, row);

  // Enable automatic follow-up sending after the first outreach.
  setAutoSendEnabled(true);

  if (stageCol > 0) {
    sh.getRange(row, stageCol).setValue('Outreach');
  }

  if (statusCol > 0) {
    const cell  = sh.getRange(row, statusCol);
    const tags  = (cell.getValue() || '')
      .toString()
      .split(',')
      .map(t => t.trim())
      .filter(Boolean);
    if (!tags.includes('Outreach')) {
      tags.push('Outreach');
      cell.setValue(tags.join(', '));
    }
  }
}


/**
 * 2) First Follow-Up: advanced threaded reply that sets To: explicitly.
 * Requires a valid Gmail thread ID for the contact.
 */
function sendFirstFollowUpForRow(email, firstName, threadId) {
  Logger.log('▶ Enter sendFirstFollowUpForRow; email=%s, firstName=%s, threadId=%s', email, firstName, threadId);

  const subject = OUTREACH_SUBJECT;
  let thread = null;
  if (threadId) {
    try {
      thread = GmailApp.getThreadById(threadId);
    } catch (err) {
      Logger.log('Failed to get thread by ID %s: %s', threadId, err);
    }
  }
  if (!thread) {
    Logger.log('Thread not found for %s; skipping first follow-up.', email);
    return;
  }
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
 * Requires a valid Gmail thread ID.
 */
function sendSecondFollowUpForRow(email, firstName, threadId) {
  Logger.log('▶ Enter sendSecondFollowUpForRow; email=%s, firstName=%s, threadId=%s', email, firstName, threadId);

  const subject = OUTREACH_SUBJECT;
  let thread = null;
  if (threadId) {
    try {
      thread = GmailApp.getThreadById(threadId);
    } catch (err) {
      Logger.log('Failed to get thread by ID %s: %s', threadId, err);
    }
  }
  if (!thread) {
    Logger.log('Thread not found for %s; skipping second follow-up.', email);
    return;
  }
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

/**
 * 4) Third Follow-Up.
 * Requires a valid Gmail thread ID.
 */
function sendThirdFollowUpForRow(email, firstName, threadId) {
  Logger.log('▶ Enter sendThirdFollowUpForRow; email=%s, firstName=%s, threadId=%s', email, firstName, threadId);

  const subject = OUTREACH_SUBJECT;
  let thread = null;
  if (threadId) {
    try {
      thread = GmailApp.getThreadById(threadId);
    } catch (err) {
      Logger.log('Failed to get thread by ID %s: %s', threadId, err);
    }
  }
  if (!thread) {
    Logger.log('Thread not found for %s; skipping third follow-up.', email);
    return;
  }
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
 * Requires a valid Gmail thread ID.
 */
function sendFourthFollowUpForRow(email, firstName, threadId) {
  Logger.log('▶ Enter sendFourthFollowUpForRow; email=%s, firstName=%s, threadId=%s', email, firstName, threadId);

  const subject = OUTREACH_SUBJECT;
  let thread = null;
  if (threadId) {
    try {
      thread = GmailApp.getThreadById(threadId);
    } catch (err) {
      Logger.log('Failed to get thread by ID %s: %s', threadId, err);
    }
  }
  if (!thread) {
    Logger.log('Thread not found for %s; skipping fourth follow-up.', email);
    return;
  }
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
 * with From:, To:, Subject:, In-Reply-To, and References headers.
 */
function buildRawMessage_(to, subject, textBody, htmlBody, inReplyTo) {
  const nl       = '\r\n';
  const boundary = '----=_Boundary_' + Date.now();

  let headers =
    `From: ${FROM_ADDRESS}` + nl +
    `To: ${to}` + nl +
    `Subject: ${subject}` + nl;

  if (inReplyTo) {
    headers +=
      `In-Reply-To: ${inReplyTo}` + nl +
      `References: ${inReplyTo}` + nl;
  }

  headers +=
    `MIME-Version: 1.0` + nl +
    `Content-Type: multipart/alternative; boundary="${boundary}"` + nl + nl;

  const body =
    `--${boundary}` + nl +
    `Content-Type: text/plain; charset="UTF-8"` + nl + nl +
    textBody + nl + nl +
    `--${boundary}` + nl +
    `Content-Type: text/html; charset="UTF-8"` + nl + nl +
    htmlBody + nl + nl +
    `--${boundary}--`;

  const msg = headers + body;
  return Utilities.base64EncodeWebSafe(msg);
}

/**
 * Helper: extracts just the email address from a "Name <email>" string.
 *
 * @param {string} from The header value.
 * @return {string} Lowercase email address.
 */
function extractEmail_(from) {
  const match = from.match(/<([^>]+)>/);
  return (match ? match[1] : from).toLowerCase();
}

/**
 * Helper: collect all sending addresses including Gmail aliases.
 *
 * @return {string[]} Lowercase addresses that belong to the account.
 */
function getMyAddresses_() {
  const aliases = GmailApp.getAliases();
  return [FROM_ADDRESS.toLowerCase()].concat(
    aliases.map(a => a.toLowerCase())
  );
}

/**
 * Helper: checks if the address belongs to the script owner or any alias.
 *
 * @param {string} addr Email address to test.
 * @return {boolean} True if the address is one of ours.
 */
function isMyAddress_(addr) {
  addr = addr.toLowerCase();
  const mine = GmailApp.getAliases()
    .map(a => a.toLowerCase())
    .concat(Session.getActiveUser().getEmail().toLowerCase(), FROM_ADDRESS.toLowerCase());
  return mine.some(a => a === addr);
}

/**
 * Determines the reply status for a thread.
 *
 * @param {GmailThread} thread Gmail thread to examine.
 * @param {string} email       Contact email address.
 * @return {string} Status: "New Response", "Replied", or "Waiting".
 */
function getLatestThreadStatus_(thread, email) {
  const messages = thread.getMessages();
  if (!messages.length) return 'Waiting';

  const contactAddr = email.toLowerCase();
  const lastAddr = extractEmail_(
    messages[messages.length - 1].getFrom()
  ).toLowerCase();
  const contactEver = messages.some(
    m => extractEmail_(m.getFrom()).toLowerCase() === contactAddr
  );

  if (lastAddr === contactAddr) {
    return 'New Response';
  }

  if (isMyAddress_(lastAddr) && contactEver) {
    return 'Replied';
  }

  return contactEver ? 'Replied' : 'Waiting';
}

/**
 * Helper to write a status value linked to the Gmail thread.
 *
 * @param {Range} cell      The Reply Status cell to update.
 * @param {string} text     Display text such as "New Response".
 * @param {string} threadId Gmail thread ID for the hyperlink.
 * @param {string} color    Background color to apply.
 */
function setReplyStatusWithLink_(cell, text, threadId, color) {
  const rich = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setLinkUrl(GMAIL_THREAD_LINK_PREFIX + threadId)
    .build();
  cell.setRichTextValue(rich).setBackground(color);
}

/**
 * Automatically send follow-up emails if contacts haven't replied.
 * Requires a stored Thread ID for each contact and skips any that are missing.
 * Intended to run daily via a time-based Apps Script trigger.
 */
function autoSendFollowUps() {
  if (!isAutoSendEnabled()) return;
  const remaining = MailApp.getRemainingDailyQuota();
  if (remaining < 10) {
    Logger.log('Daily quota low (%s emails remaining); skipping follow-ups.', remaining);
    return;
  }
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sh   = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sh) return;
  const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const firstNameCol = hdrs.indexOf('First Name') + 1;
  const lastNameCol  = hdrs.indexOf('Last Name') + 1;
  const emailCol     = hdrs.indexOf('Email') + 1;
  const statusCol    = hdrs.indexOf('Status') + 1;
  const stageCol     = hdrs.indexOf('Stage') + 1;
  const replyCol     = hdrs.indexOf('Reply Status') + 1;
  const threadIdCol  = hdrs.indexOf('Thread ID') + 1;
  if (
    firstNameCol < 1 ||
    lastNameCol < 1 ||
    emailCol < 1 ||
    statusCol < 1 ||
    stageCol < 1 ||
    replyCol < 1 ||
    threadIdCol < 1
  ) {
    throw new Error('Headers required: First Name, Last Name, Email, Status, Stage, Reply Status, Thread ID');
  }

  const numRows = sh.getLastRow() - 1;
  if (numRows < 1) return;
  const data = sh.getRange(2, 1, numRows, sh.getLastColumn()).getValues();

  data.forEach((vals, idx) => {
    const row    = idx + 2;
    const email  = vals[emailCol - 1];
    if (!email) return;
    const first  = (vals[firstNameCol - 1] || '').toString();
    const last   = (vals[lastNameCol - 1]  || '').toString();
    if (!first && !last) return;
    let status   = vals[statusCol - 1] || '';
    const storedThreadId = vals[threadIdCol - 1];
    const tags   = status.split(',').map(t => t.trim()).filter(Boolean);
    if (!tags.includes('Outreach')) return;
    if (tags.includes('Moved to DM')) return;

    let thread = null;
    if (storedThreadId) {
      try {
        thread = GmailApp.getThreadById(storedThreadId);
      } catch (err) {
        Logger.log('Failed to fetch thread %s for %s: %s', storedThreadId, email, err);
      }
    }
    if (!thread) {
      Logger.log('Thread not found for %s; skipping follow-ups.', email);
      return;
    }
    const replyCell = sh.getRange(row, replyCol);
    const threadStatus = getLatestThreadStatus_(thread, email);
    const statusColor =
      threadStatus === 'New Response' || threadStatus === 'Replied'
        ? NEW_RESPONSE_COLOR
        : null;
    setReplyStatusWithLink_(replyCell, threadStatus, thread.getId(), statusColor);

    if (threadStatus === 'New Response' || threadStatus === 'Replied') {
      if (threadStatus === 'Replied' && !tags.includes('Replied')) {
        tags.push('Replied');
      }
      const newStatus = tags.join(', ');
      if (newStatus !== status) {
        sh.getRange(row, statusCol).setValue(newStatus);
      }
      return;
    }

    const lastMsg  = thread.getMessages().pop();
    const minutesSince = (Date.now() - lastMsg.getDate().getTime()) / 60000;

    if (!tags.includes('1st Follow Up Sent') && minutesSince >= FIRST_FU_DELAY_MINUTES) {
      sendFirstFollowUpForRow(email, first, thread.getId());
      tags.push('1st Follow Up Sent');
      if (stageCol > 0) {
        sh.getRange(row, stageCol).setValue('Follow Up 1');
      }
    } else if (
      tags.includes('1st Follow Up Sent') &&
      !tags.includes('2nd Follow Up Sent') &&
      minutesSince >= SECOND_FU_DELAY_MINUTES
    ) {
      sendSecondFollowUpForRow(email, first, thread.getId());
      tags.push('2nd Follow Up Sent');
      if (stageCol > 0) {
        sh.getRange(row, stageCol).setValue('Follow Up 2');
      }
    } else if (
      tags.includes('2nd Follow Up Sent') &&
      !tags.includes('3rd Follow Up Sent') &&
      minutesSince >= THIRD_FU_DELAY_MINUTES
    ) {
      sendThirdFollowUpForRow(email, first, thread.getId());
      tags.push('3rd Follow Up Sent');
      if (stageCol > 0) {
        sh.getRange(row, stageCol).setValue('Follow Up 3');
      }
    } else if (
      tags.includes('3rd Follow Up Sent') &&
      !tags.includes('4th Follow Up Sent') &&
      minutesSince >= FOURTH_FU_DELAY_MINUTES
    ) {
      sendFourthFollowUpForRow(email, first, thread.getId());
      tags.push('4th Follow Up Sent');
      if (stageCol > 0) {
        sh.getRange(row, stageCol).setValue('Follow Up 4');
      }
    } else if (
      tags.includes('4th Follow Up Sent') &&
      !tags.includes('Moved to DM') &&
      minutesSince >= FOURTH_FU_DELAY_MINUTES
    ) {
      setReplyStatusWithLink_(
        replyCell,
        'Moved to DM',
        thread.getId(),
        MOVED_TO_DM_COLOR,
      );
      tags.push('Moved to DM');
      if (stageCol > 0) {
        sh.getRange(row, stageCol).setValue('DM');
      }
    }

    const newStatus = tags.join(', ');
    if (newStatus !== status) {
      sh.getRange(row, statusCol).setValue(newStatus);
    }
  });
}

