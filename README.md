# Kalm Follow-Up Automation

This repository contains a Google Apps Script project that automates Gmail outreach and follow‑up emails driven from a Google Sheet.

## Purpose

The script sends an initial outreach email and up to four follow‑up messages. Contacts are listed in a spreadsheet and tagged with the next action. When you update the **Status** column in the sheet, the script sends the appropriate email and schedules further follow‑ups.

## Deployment

1. Create a new Apps Script project or open an existing one attached to your spreadsheet.
2. Replace the default `Code.gs` with the contents of `code.gs` from this repository.
3. For each HTML file (`OutreachTemplate.html`, `FirstFollowUpTemplate.html`, etc.), add a new HTML template and paste in the corresponding file contents.
4. In the Apps Script editor choose **Extensions → Advanced Google services** and enable **Gmail API**.
5. Click the link to the Google Cloud Platform console and also enable the Gmail API for the project there.
6. Save and authorize the script when prompted.
7. Configure the `FROM_ALIAS` constant in `code.gs` to the Gmail alias you want
   to send from. That alias must be added as a valid **Send mail as** address in
   the Gmail settings of the account running the script.

### Configuration

The `FROM_ALIAS` constant controls which Gmail alias the script uses to send messages. Set it to one of your Gmail "Send mail as" addresses and ensure that alias is authorized for the account.

## Basic Usage

1. In your spreadsheet create columns titled **First/Last Name**, **Email**, **Status**, and **Reply Status**.
2. Install an **On edit** trigger for the `onEditTrigger` function.
3. Install a daily time‑driven trigger for `autoSendFollowUps` so unanswered threads continue to receive follow‑ups automatically.
4. Add a row for each contact and update the **Status** cell with tags such as `Outreach`, `1st Follow Up`, etc. Editing the status will send the matching email template.
5. Customize the template text and delay constants in `code.gs` as needed.
6. When a contact replies, `autoSendFollowUps` writes `New Response` to the **Reply Status** column and highlights the cell in light green. If your message is the most recent, the cell is cleared and its background reset.

With the Gmail service enabled and triggers installed, the script manages your outreach and follow‑ups directly from Gmail while updating status information in your spreadsheet.
