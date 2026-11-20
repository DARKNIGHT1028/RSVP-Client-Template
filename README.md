# RSVP Client Template

A reusable invitation and RSVP website template based on the original SeJasmin RSVP layout. The guest-facing invitation and admin dashboard are included, but the template is disconnected from the original event's Google Sheet and Apps Script deployment.

## Files
- `index.html` — invitation page and RSVP form.
- `admin.html` — password-protected RSVP dashboard with search, CSV export, print and delete.
- `Code.gs` — Google Apps Script backend.
- `README.md` — setup guide.

## Before publishing
1. Edit the event name, subtitle, date, time, venue, map location and invitation wording in `index.html`.
2. Create a new private Google Sheet for this event.
3. Open Google Apps Script, paste `Code.gs`, and replace `PASTE_YOUR_GOOGLE_SHEET_ID_HERE` with your new spreadsheet ID.
4. In Apps Script **Project Settings → Script Properties**, add `ADMIN_KEY` with a strong password for the admin dashboard.
5. Deploy the script as a Web App (**Execute as: Me**, **Who has access: Anyone**) and copy its `/exec` URL.
6. Paste that URL into the `RSVP_API` constant in `index.html` and the `API` constant in `admin.html`.
7. Enable GitHub Pages from **Settings → Pages → Deploy from a branch → main → /(root)**.

## Important
- This template does not connect to the original SeJasmin spreadsheet or Apps Script.
- Use a separate spreadsheet, Apps Script deployment and admin key for every client/event.
- The background music has been detached; add music that you have permission to use if desired.
- The countdown and calendar export require event-specific configuration before use.
- Keep the Google Sheet private. The public Web App endpoint accepts guest submissions; do not store sensitive information in this RSVP form.

## Pages
- Guest invitation: `index.html`
- Admin dashboard: `admin.html`
