# Visitor registration: how it works

The **Register Now** form on the website sends each registration to a Google Apps Script
called **Wonder Women Utsav RSVP** (in Tanishq's Google account, at script.google.com).
The script:

1. Adds the visitor to the **RSVPs** tab of the Google Sheet **Wonder Women Utsav – RSVPs**.
   The same phone number is only added once.
2. Emails them a confirmation, if they left "Send my confirmation and event updates by email" ticked.
   The email shows the Wonder Women Utsav logo, the date, time and venue, a Get Directions button,
   and a WhatsApp chat link to +91 76782 73883.

`Code.gs` in this folder is a copy of the script. The live copy is the one in Apps Script.

## Changing the script

1. Open the project at script.google.com, edit `Code.gs`, and save.
2. **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.**
   This keeps the same web-app link, so the website needs no change.
   (Creating a *new* deployment gives a new link, and the website would then need updating.)
3. Copy the new code into `rsvp-backend/Code.gs` here so the two stay the same.

## Useful functions (choose one in the dropdown, then click Run)

| Function | What it does |
|---|---|
| `setup` | Creates the RSVPs tab if it's missing, refreshes the headings while the list is empty, and sends a test confirmation to your own Gmail. |
| `removeTestRows` | Deletes any rows whose name starts with "TEST Entry". |

## Limits

A regular Gmail account can send about 100 emails a day (Google Workspace: about 1,500).
If the limit is reached, the registration is still saved and the sheet's Email status shows `Failed`.

## What's in the sheet

| Column | Meaning |
|---|---|
| Phone | Saved as text with +91. The same number is only added once. |
| Email updates OK | Whether they left the email box ticked |
| Email status | `Sent`, `Not sent (box unticked)`, or `Failed: …` with the reason |
| Source | The page they registered from |
