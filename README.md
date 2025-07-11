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
 
- `AutoSendEnabled` is a script property that controls whether follow-ups are sent
  automatically. The property is set to `TRUE` the first time an outreach email is
  sent so follow-ups start immediately. You can disable auto-sending anytime from
 **Project Settings → Script properties** by setting `AutoSendEnabled` to
 `FALSE`. To flip this value without opening settings, run the
  `toggleAutoSendEnabled` function from **Extensions → Macros** or assign it to a
  sheet button.

Toggling this value only stops follow-ups logically. The time-driven trigger
 continues to invoke `autoSendFollowUps`, which consumes an execution each day.
 Delete the trigger entirely if you need to pause scheduled runs.

All automated emails include a custom `X-Kalm-Auto` header. If you reply
manually from the same account, the absence of this header lets the script know
to stop further follow-ups automatically.
 
 ## Basic Usage
 

1. In your spreadsheet create columns titled **First Name**, **Last Name**, **Email**, **Entry Date**, **Status**, **Stage**, and **Email Link**.
 2. Install an **On edit** trigger for the `onEditTrigger` function.
 3. Install a daily time‑driven trigger for `autoSendFollowUps` so unanswered threads continue to receive follow‑ups automatically.
4. Add a row for each contact and tag the **Status** cell with `Outreach` to begin the sequence.
   Editing the status sends the matching template and updates the **Stage** column immediately.
   Follow-ups are only sent while the `Outreach` tag remains; remove it to stop further emails.
   The first outreach email automatically enables auto-sending so subsequent follow-ups are queued without extra steps.
 5. Customize the template text and delay constants in `code.gs` as needed.
6. Each run checks the time since the last message in every thread and sends the next follow‑up when due.
   The **Stage** column records which template was sent and is updated before each email is delivered.
    The **Email Link** column automatically stores a link to the Gmail thread
    when the initial outreach email is sent.
    The **Entry Date** column records the date each contact's email was added.

 
 With the Gmail service enabled and triggers installed, the script manages your outreach and follow‑ups directly from Gmail while updating status information in your spreadsheet.
 
 ### Optional: Button to Start Outreach
 
 You can place a button in your sheet that runs `startOutreachForSelectedRow` on
 the currently highlighted row:
 
 1. Insert a drawing or shape in the sheet to use as the button.
 2. Click the shape's menu (three dots) and choose **Assign script**.
 3. Enter `startOutreachForSelectedRow` and save.
 Clicking this button sends the initial outreach email, tags the **Status** cell with `Outreach`, and sets the **Stage** column to `Outreach`.
 
 ### Optional: Toggle Auto-Send
 
 Import `toggleAutoSendEnabled` under **Extensions → Macros → Import** to make it
 available from the Macros menu. You can also assign this function to another
 sheet button. Each time it runs, the `AutoSendEnabled` property switches between
 `TRUE` and `FALSE` and a dialog confirms the new state.
 This does not stop the time-driven trigger from running; delete that trigger if
 you want to halt follow-up checks completely.
