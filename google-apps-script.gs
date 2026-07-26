/**
 * eSTEAMED Academy — Website Form Handler (Google Apps Script)
 * ---------------------------------------------------------------
 * Handles submissions from the 4 website forms (contact, partnership,
 * scholarship, instructor). Paste this whole file into a new Apps Script
 * project bound to (or associated with) your Google Sheet.
 *
 * SETUP STEPS
 * 1. Go to https://sheet.new and create a Google Sheet. In it, create four
 *    tabs named EXACTLY: Contact | Partnership | Scholarship | Instructor
 *    (see SHEET_TABS below — rename there too if you use different names).
 * 2. Extensions > Apps Script. Delete the placeholder code and paste this file.
 * 3. Edit the RECIPIENT_EMAILS map below with real recipient addresses.
 * 4. Click Deploy > New deployment > type: Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Click Deploy, authorize the requested permissions, and copy the URL
 *    ending in /exec.
 * 6. Paste that URL into eSTEAMED Website.dc.html, replacing the value of
 *    GOOGLE_SCRIPT_URL near the top of the <script> block (search for
 *    "PASTE_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE").
 * 7. Whenever you change this script, you must create a NEW deployment
 *    version (Deploy > Manage deployments > Edit > New version) for the
 *    live /exec URL to pick up your changes.
 */

// ── CONFIG: edit these ──────────────────────────────────────────────

// Sheet tab names — must match tabs in your spreadsheet exactly.
const SHEET_TABS = {
  contact: 'Contact',
  partnership: 'Partnership',
  scholarship: 'Scholarship',
  instructor: 'Instructor'
};

// Recipient email per form type. Replace with real addresses.
const RECIPIENT_EMAILS = {
  contact: 'srugm25@gmail.com',
  partnership: 'srugm25@gmail.com',
  scholarship: 'srugm25@gmail.com',
  instructor: 'srugm25@gmail.com'
};

// Subject line shown in the notification email per form type.
const EMAIL_SUBJECTS = {
  contact: 'New Contact Form Submission — eSTEAMED Academy',
  partnership: 'New Partnership Inquiry — eSTEAMED Academy',
  scholarship: 'New Scholarship Application — eSTEAMED Academy',
  instructor: 'New Instructor / Team Application — eSTEAMED Academy'
};

// Honeypot field name — must match the hidden input name in the HTML forms.
const HONEYPOT_FIELD = 'hp_website';

// ── MAIN ENTRY POINT ────────────────────────────────────────────────

function doPost(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const formType = params.formType;

    // Spam check: if the honeypot field has any value, silently pretend success.
    if (params[HONEYPOT_FIELD]) {
      return jsonResponse({ result: 'success' });
    }

    // Validate form type against the four allowed values.
    if (!formType || !SHEET_TABS.hasOwnProperty(formType)) {
      return jsonResponse({ result: 'error', message: 'Invalid form type.' });
    }

    const tabName = SHEET_TABS[formType];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
    if (!sheet) {
      return jsonResponse({ result: 'error', message: 'Server configuration error.' });
    }

    // Build the field list (excluding internal/control fields), sanitized.
    const fields = [];
    for (const key in params) {
      if (key === 'formType' || key === HONEYPOT_FIELD) continue;
      fields.push({ key: key, value: sanitize(params[key]) });
    }

    // Write a row: Timestamp, then each field value in submission order.
    const timestamp = new Date();
    const headerRow = getOrCreateHeaderRow(sheet, fields);
    const row = [timestamp];
    fields.forEach(f => row.push(f.value));
    sheet.appendRow(row);

    // Send notification email.
    sendNotificationEmail(formType, fields, params.email);

    return jsonResponse({ result: 'success' });
  } catch (err) {
    // Never leak internal error details to the client.
    return jsonResponse({ result: 'error', message: 'Submission failed. Please try again later.' });
  }
}

// ── HELPERS ─────────────────────────────────────────────────────────

// Ensures the sheet has a header row matching current fields; adds one if the
// sheet is empty. (Does not reorder existing headers on later submissions —
// if you add new form fields later, add the column manually to keep alignment.)
function getOrCreateHeaderRow(sheet, fields) {
  if (sheet.getLastRow() === 0) {
    const headers = ['Timestamp'].concat(fields.map(f => f.key));
    sheet.appendRow(headers);
  }
  return true;
}

function sendNotificationEmail(formType, fields, submitterEmail) {
  const to = RECIPIENT_EMAILS[formType];
  if (!to) return;

  const subject = EMAIL_SUBJECTS[formType] || 'New Website Form Submission';
  let body = 'A new ' + formType + ' form was submitted on the eSTEAMED Academy website.\n\n';
  fields.forEach(f => {
    body += capitalize(f.key) + ': ' + f.value + '\n';
  });
  body += '\nSubmitted: ' + new Date().toLocaleString();

  const options = {};
  if (submitterEmail && isValidEmail(submitterEmail)) {
    options.replyTo = submitterEmail;
  }

  MailApp.sendEmail(to, subject, body, options);
}

function sanitize(value) {
  if (value === undefined || value === null) return '';
  // Strip characters that could be used for header/script injection in email bodies.
  return String(value).replace(/[\r\n]+/g, ' ').trim().slice(0, 2000);
}

function capitalize(s) {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
