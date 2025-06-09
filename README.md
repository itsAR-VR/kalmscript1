# Kalm Follow-Up Automation

This repository contains a Google Apps Script project that automates Gmail outreach and follow‑up emails driven from a Google Sheet.

## Purpose

The script sends an initial outreach email and up to four follow‑up messages. Contacts are listed in a spreadsheet and tagged with the next action. When you update the **Status** column in the sheet, the script sends the appropriate email and schedules further follow‑ups.

## Step-by-Step Setup

1. Create or open an Apps Script project attached to your spreadsheet.
2. Replace the default `Code.gs` with `code.gs` from this repository and create HTML templates from each `*.html` file.
3. In the Apps Script editor open **Extensions → Advanced Google services** and enable **Gmail API**, then follow the link to the Google Cloud console to enable it there as well.
4. Set the `FROM_ADDRESS` constant in `code.gs` to the Gmail address that will send your outreach messages.
5. Install an **On edit** trigger for `onEditTrigger` and a time‑driven trigger for `autoSendFollowUps`.
6. Add a drawing or button in the sheet and assign the `startOutreachForSelectedRow` function to send outreach for the active row.
7. Save and authorize the script when prompted.

### Configuration

The `FROM_ADDRESS` constant controls which Gmail address the script uses to send messages. Set it to the single account that will manage your outreach. The script checks incoming replies on this same address to stop follow‑ups automatically.
Any "Send mail as" aliases configured in Gmail are detected automatically, so replies to those addresses are also recognized.

`NEW_RESPONSE_COLOR` sets the background color applied to the **Reply Status** cell when a contact replies. The default is `red` but you can change it to any valid Sheets color name or hex value.

`AutoSendEnabled` is a script property that controls whether follow-ups are sent
automatically. Clicking the outreach button sets it to `TRUE`. You can disable
auto-sending anytime from **Project Settings → Script properties** by setting
`AutoSendEnabled` to `FALSE`.

## Basic Usage

1. In your spreadsheet create columns titled **First Name**, **Last Name**, **Email**, **Status**, and **Reply Status**.
2. Install an **On edit** trigger for the `onEditTrigger` function.
3. Install a daily time‑driven trigger for `autoSendFollowUps` so unanswered threads continue to receive follow‑ups automatically.
4. Add a row for each contact and update the **Status** cell with tags such as `Outreach`, `1st Follow Up`, etc. Editing the status will send the matching email template.
5. Customize the template text and delay constants in `code.gs` as needed.
6. Each run examines the most recent message in every thread. If the contact wrote last the **Reply Status** cell shows `New Response` in red. Once you respond it changes to `Replied`; otherwise it reads `Waiting`. The status text always links back to the Gmail thread and is refreshed even if you clear the cell. After the final follow‑up the script marks `Moved to DM`.
7. The follow-up routine searches Gmail using `in:anywhere (to:EMAIL OR from:EMAIL) subject:"SUBJECT"` so threads are matched whether messages were sent to the contact or received from them.

With the Gmail service enabled and triggers installed, the script manages your outreach and follow‑ups directly from Gmail while updating status information in your spreadsheet.

### Optional: Button to Start Outreach

You can place a button in your sheet that runs `startOutreachForSelectedRow` on
the currently highlighted row:

1. Insert a drawing or shape in the sheet to use as the button.
2. Click the shape's menu (three dots) and choose **Assign script**.
3. Enter `startOutreachForSelectedRow` and save.

Now clicking the button will send the initial outreach email for the active row
and tag its **Status** cell with `Outreach`.
