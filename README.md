# Gmail Job Submission Tracker

A private Google Sheets and Gmail workflow for tracking job submissions in one place. Label a relevant Gmail conversation and the connected Google Apps Script adds a single tracker row for that conversation.

## What it tracks

The tracker includes:

- submission date and status
- job title, designation, location, and work mode (onsite, remote, or hybrid)
- end client and vendor
- recruiter name, phone number, and email address
- rate or salary, job or requisition ID, follow-up, contact, and interview dates
- notes, the Gmail thread ID, and original email subject

## Repository files

| File | Purpose |
| --- | --- |
| `GmailJobTracker.gs` | Google Apps Script that reads labeled Gmail conversations and updates the Google Sheet. |
| `job_submissions_tracker.csv` | Header-only CSV template for importing the tracker columns into a new Google Sheet. |

## Setup

### 1. Create the Google Sheet

1. Create a new Google Sheet.
2. Choose **File → Import → Upload** and upload `job_submissions_tracker.csv`.
3. Use **Replace current sheet** when importing.
4. Note the exact name of the sheet tab at the bottom of the workbook.

### 2. Add the Apps Script

1. In the Google Sheet, choose **Extensions → Apps Script**.
2. Replace the default `Code.gs` content with the contents of `GmailJobTracker.gs`.
3. If the tab is not called `Job Submissions`, update this setting to its exact name:

   ```javascript
   const TRACKER_SHEET = 'Job Submissions';
   ```

4. Save the script.

### 3. Create the Gmail label

1. In Gmail, create a label named exactly `JobSubmissions`.
2. Apply that label to recruiter emails or job-submission conversations you want tracked.

### 4. Authorize and import

1. In Apps Script, select `importJobSubmissions` in the function menu.
2. Click **Run** and complete Google's permission flow.
3. Return to the Sheet and refresh it.

The script requires access to the bound Sheet and Gmail because it reads the conversations carrying the `JobSubmissions` label. It does not send email or modify Gmail messages.

## How it avoids duplicates

Gmail labels are applied to conversations (threads), which may contain multiple messages. The importer stores each conversation's Gmail thread ID and adds one row per thread.

If an older version of the script created several rows for one conversation, run `deduplicateExistingImports` once from Apps Script. It preserves non-Gmail/manual rows and keeps one Gmail-imported row per conversation.

## Automation

To import new labeled conversations automatically:

1. In Apps Script, click **Triggers** (the alarm-clock icon).
2. Click **+ Add Trigger**.
3. Select `importJobSubmissions` as the function.
4. Set the event source to **Time-driven** and choose an hourly interval, such as every 6 hours.
5. Save and authorize the trigger.

Google may run time-based triggers a few minutes before or after the selected interval. New rows will appear after the next trigger run.

## Notes and limitations

- The script extracts common labels such as `Job Title`, `Client`, `Vendor`, `Location`, `Rate`, and `Job ID`. Recruiter email formats vary, so review imported values before relying on them.
- The first email in each labeled conversation is used as the primary job-description source.
- Do not run `setupJobTracker` on a sheet that already contains data: it clears and rebuilds that tab. The CSV import already provides the required headers.
- Keep private resumes, candidate data, and production credentials out of this repository.

## Future improvements

- Improve extraction rules for the email formats used by frequent recruiters.
- Add a dashboard for application status, interview counts, and follow-up reminders.
- Add validation lists and conditional formatting directly in the Google Sheet.
- Add a manual review queue for messages with incomplete extracted details.
