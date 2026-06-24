# Market Data Automation

One file: **`run.py`**. Edit the `CONFIG` block at the top, run the script, review the Outlook draft.

## Prerequisites

- Windows with desktop **Outlook** logged in
- **Excel** with the Bloomberg add-in installed
- **Bloomberg Terminal** running on the same machine
- Python 3.10+

## Setup

```powershell
cd C:\path\to\carlopdf
py -3 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
py -3 -m pip install -r requirements-bloomberg.txt --index-url https://blpapi.bloomberg.com/repository/releases/python/simple/
notepad run.py
```

Open `run.py` and edit the **CONFIG** section at the top (lines ~36–140):

| Setting | What to set |
|---------|-------------|
| `email_source.folder` | Your QES folder, e.g. `"Inbox/QES"` |
| `email_source.subject` | Exact daily email subject line |
| `excel.workbook` | Full path to your `.xlsm` file |
| `excel.bloomberg_addin` | Usually `C:\blp\API\Office Tools\BloombergUI.xla` |
| `output.draft_to` | Recipient email address(es) |

## Verify & run

```powershell
py -3 run.py --check      # preflight: Outlook, Excel paths, Bloomberg/xbbg
py -3 run.py --no-poll    # live test — opens Outlook draft (nothing sent)
py -3 run.py              # production run with email poll loop
```

## Task Scheduler

Daily **4:25 PM** → `python.exe run.py` → **Run only when user is logged on**.

## CONFIG sections

- **EMAIL** — QES folder, subject, time window, chart matching
- **EXCEL** — Carlo Sectors + Desk Flows ranges, Bloomberg refresh
- **BLOOMBERG / SPTSX** — index summary, MOV by sub-industry, performers chart
- **OUTPUT** — draft recipients and subject

Optional: `--config path.json` overrides the inline CONFIG for testing.

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Email not found | Widen `received_after` / `received_before`; check `subject` |
| Wrong chart picked | Tune `chart.title_patterns` / `exclude_patterns` in CONFIG |
| Bloomberg refresh fails | Confirm `bloomberg_addin` path; open Excel once manually |
| Empty MOV columns | Tune `member_fields` in CONFIG (see comments in run.py) |
| xbbg DLL error | Install `blpapi` via `requirements-bloomberg.txt` |
