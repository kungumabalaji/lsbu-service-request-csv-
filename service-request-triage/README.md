# Service Request Triage – Prototype

A small team receives service requests in a spreadsheet. Today one person reads every
row, picks a category, writes an acknowledgement and decides which requests need
attention. This prototype does that sorting automatically and consistently, and leaves
the decisions that need judgement to a person.

## What it does

For every request in the CSV file, the script:

| Step (from the brief) | What the script does |
|---|---|
| 1. Read the requests | Reads the CSV and checks that all 8 required columns exist |
| 2. Check required information | Flags blank fields, invalid priority, channel, status or request type, invalid or future dates (UK DD/MM/YYYY), badly formatted or duplicate IDs, and summaries that are too short |
| 3. Categorise with simple rules | Matches whole-word keywords in the **Request Summary** against `rules.csv` (first match wins), records the keywords that matched, any other categories that also matched, and a confidence level |
| 4. Suggest an acknowledgement | Fills a fixed template with the ID, category and priority. **It is never sent automatically.** No draft is made for invalid rows or for requests that are already completed |
| 5. Identify requests for human review | Flags each request that needs a person and records **why** (see below) |
| 6. Create an output | Writes an Excel report with **Results**, **Summary** and **Rules** sheets, and prints the summary |

### When a person must review a request

- Any data issue
- High priority
- Access & Permissions requests (access must be approved by a person)
- The summary mentions external or third-party access, finance or payments, personal or sensitive data, or passwords (keyword lists are in `rules.csv`)
- No category rule matched, or the rules disagree with the Request Type the requester chose
- The summary starts with a spreadsheet formula character (`= + - @`)

Requests that are already Completed, Closed or Cancelled are not triaged again.

## Folder contents

```
triage.py                 the prototype
rules.csv                 category and review keywords (editable in Excel)
requirements.txt          openpyxl, pytest
data/requests.csv         the 10 synthetic requests supplied with the task
data/test_requests.csv    12 extra synthetic rows with errors and edge cases
tests/test_triage.py      the three test scenarios
output/triage_output.xlsx result for data/requests.csv
output/test_output.xlsx   result for data/test_requests.csv
evidence/                 test results and console output from each run
docs/                     workflow diagrams (PNG, plus both in workflow_design.pdf)
```

## Workflow design

![End-to-end workflow](docs/workflow_end_to_end.png)

![Script logic](docs/workflow_script_logic.png)

## How to run

Requires Python 3.9 or later.

```bash
pip install -r requirements.txt

python triage.py data/requests.csv output/triage_output.xlsx
python triage.py data/test_requests.csv output/test_output.xlsx

python -m pytest -v
```

On Windows, use `py` instead of `python`.

## Output (Results sheet)

The 8 original columns are kept unchanged, and these are added:

| Column | Meaning |
|---|---|
| Validation Status / Validation Issues | Valid or Invalid, and every problem found |
| Assigned Category | Category from the keyword rules, or Uncategorised |
| Matched Keywords / Also Matched | Why that category was chosen, and any other categories that also matched |
| Category Confidence | High: one clear match that agrees with the Request Type. Medium: agrees, but other categories also matched. Low: no match, or a disagreement |
| Suggested Acknowledgement | The draft reply for a person to check and send |
| Needs Human Review / Review Reasons | Yes or No, and why |

Rows with data issues are shaded red. Rows that need review are shaded amber.

**Result on the supplied data:** all 10 rows are valid and every rule-based category
agrees with the submitted Request Type. 6 requests need review (REQ-001, 002, 004, 006,
007 and 009) and 4 need no manual triage. 9 acknowledgements are drafted; REQ-010 is
skipped because it is already completed.

## Changing the rules

Open `rules.csv` in Excel. Each row is a rule:

- `category` rows: the category name and its keywords, separated by `;`. Rows are checked from top to bottom, and the first match wins.
- `review` rows: the reason a person must review the request, and the keywords that trigger it.

No code changes are needed.

## Tests

| Test | What it proves |
|---|---|
| 1 – Supplied data | All 10 rows are valid and correctly categorised; exactly the expected 6 are flagged, with the right reasons; REQ-001's ambiguity (access and dashboard) is recorded; the completed request gets no acknowledgement; the Excel report has all 3 sheets |
| 2 – Invalid data (7 cases) | Missing summary, priority "Urgent", 31/02/2026, duplicate ID, `REQ104`, a future date, and a role made only of spaces are each caught, flagged, and never acknowledged |
| 3 – Edge cases | No keyword match leads to Uncategorised and review; a rule/type disagreement is flagged and the category is left out of the acknowledgement; mixed case is tidied; 13/10/2026 is read as a UK date; multiple reasons are all recorded; a formula in a summary is stored as plain text; a file with missing columns is rejected with a clear message |

The test results are in `evidence/pytest_results.txt` (9 passed: 3 scenarios, with test 2 run for 7 inputs).
The tests were also checked by deliberately breaking the code (US date format, no formula
protection, removing the High-priority rule); each break was caught by a failing test.

## Design decisions

- **Keyword rules, not an LLM, for the core.** The brief asks for simple rules. Rules are
  predictable, explainable, free, and keep all request text inside the organisation.
- **Suggested, not sent.** A person checks every acknowledgement before it goes out.
- **Doesn't guess.** If no rule matches, or the rules disagree with the requester, a person decides.
- **Rules kept in a separate file**, so the team can maintain them without a developer.

## Limitations

- Keyword rules miss requests that use different wording (for example "can't log on")
  and can be misled by words used in another sense. They need tuning as real requests arrive.
- The prototype has been tested only on small synthetic data.
- The CSV has no requester name or email, so acknowledgements can't be sent automatically yet.
- It runs on demand. There is no scheduling or audit log of who acted on each flag.

## Security considerations

- Synthetic data only; no real personal or university information.
- No network calls. Request text never leaves the machine.
- Spreadsheet formula injection is neutralised in the output file.
- Access, external, finance and personal-data requests always go to a person.
- In production, the input and output files should sit in a restricted SharePoint or
  OneDrive location, and any real request data should be handled under UK GDPR and the
  university's data protection policy.

## How this could run in production (Microsoft 365)

1. Requests from email, Teams and the service portal land in a SharePoint list or spreadsheet.
2. Power Automate runs the same rules when a request is added (or on a schedule).
3. Routine requests are acknowledged; flagged requests post a Teams message to the right
   reviewer (IT Security, Finance or the team lead), with the reason.
4. A dashboard shows the queue, ageing and review reasons.

**Possible future improvement:** when no keyword rule matches, an approved AI assistant
could *suggest* a category. A person would still confirm it, and the suggestion would be
logged to improve the rules.
