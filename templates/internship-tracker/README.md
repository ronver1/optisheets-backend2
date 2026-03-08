# Internship Tracker — Template Docs

AI-powered recruiting coach built on top of the OptiSheets credit system.
Given the student's application tracker data, it returns prioritized, actionable recruiting recommendations.

---

## What the AI does

- Flags applications that need an urgent follow-up
- Gives interview prep tips for stages that are currently active
- Identifies pipeline weak spots (low response rate, stalled stages)
- Suggests concrete next steps

Output is plain text, ≤ 400 words, with short labeled sections.

---

## `userData` schema

```js
{
  // Required
  applications: [
    {
      company:     "Google",                  // string — company name
      role:        "SWE Intern",              // string — role / position title
      status:      "Phone Screen",            // string — see valid statuses below
      appliedDate: "2026-01-15",              // string — ISO date (YYYY-MM-DD), optional
      notes:       "Referred by Jane Doe",    // string — free text, truncated at 80 chars, optional
    },
    // ... more applications
  ],

  // Optional context
  targetRole:   "Software Engineering Intern", // defaults to "internship"
  targetSeason: "Summer 2026",                 // e.g. "Summer 2026", "Fall 2025"
  targetCount:  3,                             // number of offers the student wants
}
```

### Valid `status` values (ranked by priority in the prompt)

| Status | Meaning |
|---|---|
| `Offer` | Received an offer |
| `Final Round` | Final interview scheduled/completed |
| `Onsite` | Onsite / virtual onsite scheduled |
| `Technical Screen` | Coding/technical interview stage |
| `Phone Screen` | Recruiter or hiring manager call |
| `Applied` | Application submitted, no response yet |
| `Rejected` | Rejected at any stage |
| `Withdrawn` | Student withdrew |

Unknown statuses are accepted but sorted to the end.

---

## Token budget

`buildPrompt` caps the prompt to **~1000 input tokens** (~4000 chars):

- System prompt: ~70 tokens (fixed)
- User preamble + footer: ~50 tokens (fixed)
- Application rows: remaining budget (~880 tokens)
  - Active-stage applications are sorted first and shown first
  - Applications are dropped (oldest/lowest-priority first) if the budget is exceeded
  - Notes are hard-truncated at 80 chars per row

---

## How the Google Apps Script should call the backend

### 1. Collect data from the sheet

```javascript
function collectUserData(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; // ["Company", "Role", "Status", "Applied Date", "Notes"]
  const applications = data.slice(1)
    .filter(row => row[0]) // skip blank rows
    .map(row => ({
      company:     row[0],
      role:        row[1],
      status:      row[2],
      appliedDate: row[3] ? Utilities.formatDate(new Date(row[3]), 'UTC', 'yyyy-MM-dd') : undefined,
      notes:       row[4] || undefined,
    }));

  return {
    applications,
    targetRole:   'Software Engineering Intern', // or read from a named cell
    targetSeason: 'Summer 2026',
    targetCount:  3,
  };
}
```

### 2. Call `/api/get-recommendations`

The backend expects:
- `private_key` — the license key stored in the sheet's script properties
- `template_id` — `"internship-tracker"`
- `user_data` — the `userData` object (see schema above)
- `system_prompt` — the static system prompt string (exported as `SYSTEM_PROMPT` from `prompt.js`)

```javascript
function getRecommendations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tracker');
  const privateKey = PropertiesService.getScriptProperties().getProperty('OPTISHEETS_KEY');
  const userData = collectUserData(sheet);

  const response = UrlFetchApp.fetch('https://your-backend.vercel.app/api/get-recommendations', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      private_key:   privateKey,
      template_id:   'internship-tracker',
      user_data:     userData,
      // The backend's aiRequestHandler uses this as the OpenAI system message.
      // Copy the SYSTEM_PROMPT string from templates/internship-tracker/prompt.js here,
      // or fetch it from a config endpoint if you prefer to keep it server-managed.
      system_prompt: 'You are a recruiting coach helping college students land internships. ...',
    }),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());
  if (!result.success) {
    SpreadsheetApp.getUi().alert('Error: ' + result.error);
    return;
  }

  // Display the recommendations
  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recommendations')
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Recommendations');
  outputSheet.clearContents();
  outputSheet.getRange('A1').setValue(result.output);
  outputSheet.getRange('A1').setWrap(true);
  SpreadsheetApp.getUi().alert('Done! ' + result.remaining_credits + ' credits remaining.');
}
```

### 3. (Alternative) Server-side prompt building

If the backend is extended to call `buildPrompt(user_data)` server-side (e.g. a future
template-aware route), the Apps Script only needs to send `private_key`, `template_id`,
and `user_data` — the backend handles prompt construction and token capping automatically.

---

## Notes

- Cached responses are returned instantly without spending a credit (same `user_data` = same hash).
- The `notes` field is truncated at 80 chars in the prompt — keep notes short for best results.
- If the student has many applications, add the most important ones first in the sheet;
  the prompt will prioritize active stages but may drop lower-priority rows if over budget.
