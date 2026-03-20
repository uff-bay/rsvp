# GitHub Pages + Google Sheets RSVP starter

This starter uses:

- **GitHub Pages** for the public RSVP form UI.
- **Google Sheets + Google Apps Script** for the database, business logic, email sending, cancellation flow, and organizer controls.

What it includes:

- guaranteed-capacity vs waitlist logic
- automatic promotion from the waitlist when someone withdraws
- automatic promotion when the organizer increases capacity
- emailed self-cancel link
- a two-click cancel flow: click the email link, then click the cancel button
- a Google Sheet that non-programmers can read and manage

---

## Files

- `index.html` — public RSVP page for GitHub Pages
- `styles.css` — styles for the public page
- `config.js` — set your Apps Script `/exec` URL here
- `app.js` — loads event status and submits the RSVP form
- `google-apps-script/Code.gs` — Apps Script backend code

---

## Recommended setup flow

## 1) Create the Google Sheet

1. Create a new Google Sheet.
2. Name it something like `Tree Planting RSVP`.
3. Open **Extensions → Apps Script**.
4. Replace the default code with the contents of `google-apps-script/Code.gs`.
5. Save the project.
6. Reload the Google Sheet.
7. Open the new **Event RSVP** menu and click **Set up / repair workbook**.

That creates these tabs:

- **Settings**
- **Dashboard**
- **RSVPs**

## 2) Fill in the Settings sheet

In the `Settings` sheet, update these values:

- `event_title`
- `event_date_display`
- `event_location`
- `guaranteed_capacity`
- `coordinator_email`
- `public_form_url` (your GitHub Pages URL)
- `waiver_url_a`
- `waiver_url_b`

Leave `web_app_url` blank for now.

## 3) Deploy the Apps Script as a web app

From the Apps Script editor:

1. Click **Deploy → New deployment**.
2. Choose **Web app**.
3. For execution, choose the option that runs the app as **you / the script owner**.
4. For access, choose a setting that lets your event attendees use the form publicly.
5. Deploy it.
6. Copy the `/exec` URL.
7. Paste that URL into the `web_app_url` row in the `Settings` sheet.

Important:

- Use the **`/exec`** URL for production, not the `/dev` test URL.
- If your Google account does not offer a public access option, anonymous public RSVP will not work from GitHub Pages.

## 4) Publish the GitHub Pages frontend

1. Put `index.html`, `styles.css`, `config.js`, and `app.js` in a GitHub repository.
2. Edit `config.js` and paste your Apps Script `/exec` URL into `webAppUrl`.
3. Turn on GitHub Pages for the repo.
4. Copy the GitHub Pages URL.
5. Paste it into the `public_form_url` row in the `Settings` sheet.

If you already pasted a placeholder in `public_form_url`, update it now with the real one.

---

## How the cancel flow works

When a person RSVPs:

1. The Apps Script backend stores the RSVP in the `RSVPs` tab.
2. It sends an email to the attendee.
3. That email contains a private cancellation link.
4. The attendee clicks the link.
5. They land on a confirmation page.
6. They click **Yes, cancel my RSVP**.

That second click is the actual cancellation.

If the cancelled attendee had a confirmed spot, the backend automatically promotes the first person on the waitlist.

---

## How the organizer increases capacity

Open the Google Sheet and use:

**Event RSVP → Increase guaranteed capacity...**

Enter how many *additional* spots you want to add.

Example:

- current capacity = 40
- enter `15`
- new capacity = 55

The script then automatically promotes as many waitlisted people as needed to fill the newly opened guaranteed spots.

---

## How organizers can read the data

### `Dashboard` tab

This shows:

- guaranteed capacity
- confirmed count
- waitlist count
- withdrawn count
- remaining guaranteed spots

### `RSVPs` tab

Each RSVP row stores:

- attendee details
- current status (`confirmed`, `waitlisted`, `withdrawn`)
- waitlist position
- confirmation/promotion/withdrawal timestamps
- last email status

---

## Notes

- The public page loads live availability from Apps Script.
- The actual form submission goes to the Apps Script web app.
- After submit, the attendee sees a server-generated confirmation page.
- The script prevents exact duplicate active RSVPs for the same first name, last name, and email address.

---

## Suggested first test

Set `guaranteed_capacity` to `2`, then test this sequence:

1. Submit RSVP A → should be confirmed
2. Submit RSVP B → should be confirmed
3. Submit RSVP C → should be waitlisted as #1
4. Cancel RSVP A from the email link → RSVP C should be promoted automatically
5. Increase capacity by `1` → the next waitlisted person should be promoted automatically

