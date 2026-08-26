/**
 * Gmail Job Tracker
 *
 * 1. Create a blank Google Sheet and import job_submissions_tracker.csv.
 * 2. Open Extensions > Apps Script, replace the starter code with this file,
 *    save, and run setupJobTracker once.
 * 3. In Gmail, apply the label "JobSubmissions" to job-related emails.
 * 4. Run importJobSubmissions, or create a time-driven trigger for it.
 */

const TRACKER_SHEET = 'Job Submissions';
const GMAIL_LABEL = 'JobSubmissions';

const HEADERS = [
  'Submission Date', 'Status', 'Job Title', 'Designation', 'Work Mode',
  'End Client', 'Vendor', 'Recruiter Name', 'Recruiter Phone',
  'Recruiter Email', 'Location', 'Rate / Salary', 'Job ID', 'Source',
  'Follow-up Date', 'Last Contact Date', 'Interview Date', 'Notes',
  'Gmail Thread ID', 'Email Subject'
];

function setupJobTracker() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) sheet = ss.insertSheet(TRACKER_SHEET);
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setBackground('#17365D')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setWrap(true);
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  ['O', 'P', 'Q'].forEach(col => sheet.getRange(`${col}:${col}`).setNumberFormat('yyyy-mm-dd'));
  sheet.getRange('B2:B1000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Applied', 'Submitted', 'Interviewing', 'Offer', 'Rejected', 'Closed'], true)
      .build()
  );
  sheet.autoResizeColumns(1, HEADERS.length);
  sheet.setColumnWidth(18, 280);
  sheet.getRange(1, 1, 1000, HEADERS.length).createFilter();
}

function importJobSubmissions() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) throw new Error('Run setupJobTracker first.');

  const label = GmailApp.getUserLabelByName(GMAIL_LABEL);
  if (!label) throw new Error(`Create and apply the Gmail label "${GMAIL_LABEL}" first.`);

  const existing = new Set(
    sheet.getLastRow() > 1
      ? sheet.getRange(2, 19, sheet.getLastRow() - 1, 1).getValues().flat().filter(String)
      : []
  );
  const rows = [];
  label.getThreads().forEach(thread => {
    // Gmail labels conversations (threads), not just individual emails. Use the
    // first message as the job-description source and the thread ID as the key,
    // so each conversation is represented by one tracker row.
    if (existing.has(thread.getId())) return;
    const message = thread.getMessages()[0];
    const parsed = parseJobEmail_(message);
    rows.push([
      message.getDate(), 'Submitted', parsed.jobTitle, parsed.designation,
      parsed.workMode, parsed.endClient, parsed.vendor, parsed.recruiterName,
      parsed.phone, parsed.recruiterEmail, parsed.location, parsed.rate,
      parsed.jobId, 'Gmail', '', message.getDate(), '', parsed.notes,
      thread.getId(), message.getSubject()
    ]);
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
}

/** Run once to collapse already-imported message-level duplicates into one row per conversation. */
function deduplicateExistingImports() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TRACKER_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;

  sheet.getRange(1, 19).setValue('Gmail Thread ID');
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  const keptThreads = new Set();
  const duplicateRows = [];

  rows.forEach((row, index) => {
    if (row[13] !== 'Gmail' || !row[18]) return; // Preserve manually entered rows.
    try {
      const threadId = GmailApp.getMessageById(row[18]).getThread().getId();
      const sheetRow = index + 2;
      if (keptThreads.has(threadId)) {
        duplicateRows.push(sheetRow);
      } else {
        keptThreads.add(threadId);
        sheet.getRange(sheetRow, 19).setValue(threadId);
      }
    } catch (error) {
      // Leave any row untouched if its original Gmail message is unavailable.
    }
  });

  duplicateRows.reverse().forEach(row => sheet.deleteRow(row));
}

function parseJobEmail_(message) {
  const subject = message.getSubject();
  const body = message.getPlainBody();
  const text = `${subject}\n${body}`;
  const sender = message.getFrom();
  const senderEmail = (sender.match(/<([^>]+)>/) || [, sender])[1].trim();
  const senderName = sender.replace(/\s*<[^>]+>/, '').replace(/["']/g, '').trim();
  const value = (labels) => {
    for (const label of labels) {
      const m = text.match(new RegExp(`(?:^|\\n)\\s*${label}\\s*[:\\-]\\s*([^\\n]+)`, 'i'));
      if (m) return m[1].trim();
    }
    return '';
  };
  const phone = (text.match(/(?:\+?1[.\-\s]?)?(?:\(?\d{3}\)?[.\-\s]?)\d{3}[.\-\s]?\d{4}/) || [''])[0];
  const workMode = /\b(remote|wfh)\b/i.test(text) ? 'Remote' : /\bhybrid\b/i.test(text) ? 'Hybrid' : /\bon[ -]?site\b/i.test(text) ? 'Onsite' : '';
  return {
    jobTitle: value(['Job Title', 'Position', 'Role']) || subject,
    designation: value(['Designation', 'Title']),
    workMode,
    endClient: value(['End Client', 'Client']),
    vendor: value(['Vendor', 'Implementation Partner']),
    recruiterName: value(['Recruiter', 'Contact Name']) || senderName,
    recruiterEmail: value(['Email', 'E-mail']) || senderEmail,
    phone,
    location: value(['Location']),
    rate: value(['Rate', 'Salary', 'Compensation']),
    jobId: value(['Job ID', 'Req ID', 'Requisition ID']),
    notes: 'Imported from Gmail; review fields extracted from email.'
  };
}
