# RSVP setup: Wonder Women Utsav

The website form sends each RSVP to a small Google Apps Script attached to a Google Sheet.
The script saves the guest and sends confirmations.

## Part 1: Guest list + email confirmations (about 10 minutes)

1. Sign in to the Google account that should own the guest list and send the emails.
2. Create a new Google Sheet and name it **Wonder Women Utsav – RSVPs**.
3. In the Sheet, open **Extensions → Apps Script**.
4. Delete everything in `Code.gs` and paste in the whole of `rsvp-backend/Code.gs` from this folder. Click **Save**.
5. In the function dropdown at the top, choose **setup**, then click **Run**.
   - Google asks for permission. Choose your account, then **Advanced → Go to (project) (unsafe)** → **Allow**.
     (It says "unsafe" only because you wrote the script yourself and Google hasn't reviewed it.)
   - This creates the **RSVPs** tab and sends you a test confirmation email. Check your inbox.
6. Click **Deploy → New deployment**. Click the gear icon next to "Select type" and pick **Web app**.
   - Description: `RSVP`
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, then copy the **Web app URL** (it ends in `/exec`).
7. Send that URL to Claude, or paste it into `index.html` where it says
   `PASTE_YOUR_APPS_SCRIPT_URL_HERE`, then push the site.

**Changing the script later:** use **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**.
That keeps the same URL. (Making a *new* deployment gives a new URL, and the website would need updating.)

**Email limits:** a regular Gmail account can send about 100 emails a day. Google Workspace accounts get about 1,500.
If you run over the limit, the RSVP is still saved. Only the email is skipped, and the sheet shows `Failed`.

## Part 2: WhatsApp confirmations (a few days, mostly waiting on Meta)

Until this is done, every RSVP is still saved and the sheet's WhatsApp column shows `Pending`.
Once it's switched on, run `sendPendingWhatsApp` once and those people get their WhatsApp confirmation too.

1. **Get a phone number for the business.** It must NOT already be on the WhatsApp or WhatsApp Business app.
   (A new SIM is simplest. You can also delete WhatsApp from an existing number first.)
2. Go to **developers.facebook.com → My Apps → Create app**, choose **Business** type, and add **WhatsApp**.
   This creates (or links) a Meta Business account.
3. In **WhatsApp → API Setup**, add your business phone number and verify it by OTP.
   Copy the **Phone number ID** shown there.
4. **Create the message template.** Go to WhatsApp Manager → **Message templates → Create template**:
   - Category: **Utility**
   - Name: `rsvp_confirmation`
   - Language: **English**
   - Body:
     ```
     Hi {{1}}, your visit to Wonder Women Utsav is confirmed! 🎉

     📅 Saturday, 31 October 2026, 10 AM – 8 PM
     📍 Sunder Bhawan, Community Centre, Sunder Vihar, New Delhi

     Entry is free. See you there!
     ```
   - Sample value: `{{1}}` = `Priya`
   - Submit. Approval usually takes from a few minutes to a day.
5. **Get a permanent access token.** The "temporary token" on the API Setup page expires in 24 hours, so don't use it.
   In **Business Settings → Users → System users**, add a system user (Admin) and assign it your app and WhatsApp account.
   Then click **Generate new token** with the permissions `whatsapp_business_messaging` and `whatsapp_business_management`.
6. **Add a payment method** in WhatsApp Manager. Utility messages cost a small amount each (check Meta's current India rate card).
7. **Switch it on in the script.** In Apps Script, go to **Project Settings (⚙️) → Script properties → Add property**:

   | Property | Value |
   |---|---|
   | `WA_TOKEN` | the permanent token |
   | `WA_PHONE_NUMBER_ID` | the Phone number ID |
   | `WA_TEMPLATE` | `rsvp_confirmation` |
   | `WA_LANG` | `en` |

8. Choose **sendPendingWhatsApp** in the function dropdown and click **Run**. Everyone who registered earlier gets their message.

New Meta business accounts can message only a limited number of people per day at first.
The limit goes up automatically as you send more messages, and faster once the business is verified.

**Prefer an easier dashboard?** Interakt, AiSensy and WATI handle steps 2–6 for a monthly fee.
The script would need a small change to call their API instead. Ask Claude.

## What's in the sheet

| Column | Meaning |
|---|---|
| Phone | The same phone number is only added once. |
| WhatsApp OK | Whether they ticked the WhatsApp consent box |
| Email status / WhatsApp status | `Sent`, `Pending…`, `No consent`, or `Failed (…)` with the reason |
