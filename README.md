# OriginCrest outreach agent

One agent with three workflows. It runs on your Windows PC, reads and writes the two
OneDrive workbooks directly, uses Claude (with web search) for research and writing, and
saves every email as an **IMAP draft** for you to review. It never sends email.

| # | Workflow | When | What it does |
|---|----------|------|--------------|
| 1 | `discover` | Weekly (Mon 6:30) | Searches the web for Australian specialty roasters and for forwarders / customs brokers / 3PLs that handle green and roasted coffee. Skips anything already in either workbook, then adds new rows with status **New - to review**: Adelaide/SA businesses go on their own tab. |
| 2 | `leads` | Every 2 h, 7:00-19:00 (and right after discovery) | Finds new roaster rows, researches each one (what they roast, size, sourcing, location, current Indonesian coffees), scores fit 0-100 for Sumatra / Sulawesi / Java / Flores, and saves a personalised intro email as a draft. |
| 3 | `followup` | Weekdays 7:30 | Reads the RFQ workbook, finds **Open** rows with no reply after 5 days (quotes) or 10 days (leads), checks your mailbox for replies first, then drafts a threaded follow-up and updates **Last contacted** and **Next action date**. |

The spreadsheets stay where they are, so your existing OneDrive → Google Drive backup keeps working.

## Guardrails

- **Drafts only.** Emails are saved to your Drafts folder. Nothing goes out until you press Send.
- **Discovery rows are quarantined.** New rows get status `New - to review`, `Added by = OriginCrest Agent`
  and a `Source URL`, on their own tabs (`Roasters - Adelaide`, `Roasters - Australia`,
  `Freight-3PL - Adelaide`, `Freight-3PL - Australia`), in both workbooks.
- **Only publicly listed business contacts.** Claude is told never to guess or build email addresses.
  The agent then loads the source page (and the company's contact page) and **keeps an email, phone or
  contact name only if it actually appears there**. If the page says "no unsolicited email", the email
  is left blank. Anything dropped is explained in `Agent notes`.
- **Spam Act 2003.** Every lead email gets a footer, added by code rather than by the model, with your name,
  business, address/phone and an unsubscribe instruction ("reply 'unsubscribe'"), plus a
  `List-Unsubscribe` header. Unsubscribe replies are detected on the next weekday run, the row is set to
  `Unsubscribed - do not contact`, and the address goes on a suppression list that every workflow checks.
  The agent won't draft anything until your sender details in `config.yaml` are filled in.
- **Caps per run:** 30 new rows (10 per search segment), 10 companies researched, 8 intro drafts,
  15 follow-up drafts, max 3 follow-ups per contact. All adjustable.
- **Your data is safe:** a timestamped backup of each workbook is made before every save, nothing is written
  while a workbook is open in Excel, and your own columns (e.g. `Notes`) are never overwritten. The agent
  writes its comments to its own `Agent notes` column.
- **Leaves your contacts alone:** workflow 2 only touches rows whose status is blank, `New`, `New - to review`,
  `Reviewed`, `Approved`, `To contact`, `Lead` or `Prospect`. Rows like `Customer` or `In discussion`
  are never emailed. Before drafting an intro it also checks your Sent folder: if you've already emailed
  them, the row is set to `Open` instead.

> Not legal advice: the Spam Act lets you email a business address that is conspicuously published,
> if the message relates to their business and the page doesn't say they don't want it. The Act also
> restricts using "address-harvesting" software to build mailing lists. That's why this agent keeps a
> person in the loop: you review every row and every draft, it researches companies one at a time, it records
> where each address was published, and it sends nothing in bulk. See ACMA's Spam Act guidance if unsure.

## How the statuses move

```
discover ─► New - to review ─► (leads) ─┬─► Intro drafted - to review ─(you send it)─► Open
                                        ├─► Researched - low fit        (set Approved to draft anyway)
                                        └─► Researched - no public email

Open ─(followup, no reply after 5/10 days)─► follow-up draft, Last contacted = today, Next action date = +5/+10
Open ─(reply found)──────────────────────► Replied - review      (Last reply, Next action date = today)
Open ─(unsubscribe reply)────────────────► Unsubscribed - do not contact
Open ─(bounce)───────────────────────────► Bounced - check email
```

- Want to vet discovered rows before any intro is drafted? Set `leads.draft_for_unreviewed: false`. The agent
  then researches and scores them, and drafts only after you change the status to `Reviewed` or `Approved`.
- `Approved` also overrides a low fit score.
- RFQs you send yourself: set the row to `Open` with a `Last contacted` date (or just send from this mailbox:
  the agent picks the date up from your Sent folder) and it will be chased after 5 days.
- Quote vs lead: taken from a `Record type` column if you have one, otherwise roaster tabs are leads and
  everything else in the RFQ workbook is a quote.

## Columns

Existing columns are matched by name (e.g. `Company Name`, `Contact Person`, `Email Address`, `Date Contacted`),
and the header row doesn't have to be row 1. Missing columns are added at the right-hand end of the header row,
styled like your existing headers, and Excel tables and filters are stretched to include them.

| Workflow | Columns it may add |
|----------|--------------------|
| 1 | Company, Category, Contact name, Role, Email, Phone, Website, City, State, Services / offering, Source URL, Status, Date added, Added by, Record type, Agent notes |
| 2 (roaster tabs) | What they roast, Size estimate, Sourcing, Current Indonesian coffees, Fit score (0-100), Best-fit origins, Fit rationale, Researched on, Intro draft created, Intro draft subject |
| 3 | Last contacted, Next action date, Last reply, Follow-ups drafted, Agent notes |

## Setup (Windows, about 20 minutes)

1. **Python 3.12+** from <https://www.python.org/downloads/>. During install, tick *Add python.exe to PATH*.
2. **Get this folder onto the PC.** Use `git clone` or download the ZIP from GitHub and extract it, e.g. to
   `C:\Users\swami\OriginCrestAgent`. A local folder is better than OneDrive for the code.
3. **Install** in PowerShell, from that folder:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\setup.ps1
   ```
4. **Anthropic API key**: create one at <https://platform.claude.com/settings/keys> and put it in `.env`
   as `ANTHROPIC_API_KEY=...`.
5. **Email (IMAP)**: for Gmail, turn on 2-Step Verification, then create an *App password*
   (Google Account → Security → App passwords) and put it in `.env` as `IMAP_PASSWORD=...`.
   IMAP is enabled by default on Gmail. For Outlook/Microsoft 365, set `email.imap_host` to
   `outlook.office365.com` (your tenant must allow IMAP with an app password).
6. **Edit `config.yaml`**:
   - `email.username`: your email address.
   - `sender`: name, business address (or phone), sending email, ABN if you want it shown. These go in the
     signature and the Spam Act footer. `sender.email` should be the mailbox above or an alias of it.
   - `offer`: what the emails may say about OriginCrest (summary, your Sumatra/Sulawesi/Java/Flores lots,
     facts such as sample sizes or where stock is held). The agent states nothing beyond this.
   - The workbook paths are already set to your two files.
7. **Check everything**:
   ```powershell
   scripts\run.cmd check --api
   ```
   This lists every tab, the header row it found and how each column was matched. Fix anything marked `[X]`.
8. **First run by hand**, with Excel closed:
   ```powershell
   scripts\run.cmd init-sheets          # creates the 4 discovery tabs in both workbooks
   scripts\run.cmd discover --dry-run   # searches and reports, writes nothing
   scripts\run.cmd discover             # real run, then researches + drafts intros
   scripts\run.cmd followup --dry-run   # previews follow-ups as .eml files
   ```
9. **Schedule it**:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\install_tasks.ps1
   ```
   Change times with e.g. `-DiscoverDay Sunday -DiscoverTime 20:00 -FollowupTime 07:00`. The tasks run as
   you while you're logged in, and catch up after the PC was asleep. Remove them with
   `scripts\uninstall_tasks.ps1`.

## Your daily routine

1. Open your **Drafts** folder. Every draft is a complete email to one contact: edit if you like, then
   **Send**, or delete it.
2. Glance at the workbook: `Replied - review` rows need your answer; `New - to review` rows are new finds to
   check; `Agent notes` explains what the agent did and why.
3. Each run writes a short report to `%LOCALAPPDATA%\OriginCrestAgent\reports` (`latest_followup.md` etc.)
   listing drafts, replies, flagged rows, anything skipped, and the Claude tokens / web searches used.

## Good to know

- **Excel open?** The agent never writes to a workbook that's open. It waits up to 5 minutes, then keeps
  the changes (new rows are queued; research and drafts are remembered) and writes them on the next run.
  Close the workbooks overnight and the early-morning runs will always get through.
- **No double drafts.** Research, drafts and follow-ups are recorded in a local database
  (`%LOCALAPPDATA%\OriginCrestAgent\state.db`), so re-runs never draft the same email twice. A follow-up isn't
  drafted while an earlier draft to the same address is still waiting unsent in Drafts.
- **Deleted a junk discovery row?** It won't be re-added: the agent remembers what it found before.
- **Same company in both workbooks** is researched once and drafted once; both rows are updated.
- **Workflow 3 reads only the RFQ workbook**, as requested. Add `master` to `followup.workbooks` if you want
  its Open rows chased too (one follow-up per company either way).
- **Excel features:** the workbooks are edited with `openpyxl`. Data, formatting, formulas, tables, filters and
  simple dropdowns are kept. Some features may not survive a save: dropdowns whose list comes from another
  sheet, slicers, form controls, shapes and external links. `check` warns you if a workbook uses any of them,
  and so does the run report. If anything goes wrong, restore the copy from
  `%LOCALAPPDATA%\OriginCrestAgent\backups`.
- **Cost control:** each run report shows tokens and web searches used; compare with your Claude Console usage.
  To spend less, lower the caps or `*_searches_*` settings, or set `claude.model: claude-sonnet-5-5`.
- **Logs:** `%LOCALAPPDATA%\OriginCrestAgent\logs\agent.log`. Add `-v` for more detail on screen.

## Commands

```
scripts\run.cmd check [--api]
scripts\run.cmd init-sheets [--dry-run]
scripts\run.cmd discover [--dry-run] [--max-rows N] [--segments roasters_adelaide,freight_adelaide] [--no-leads]
scripts\run.cmd leads [--dry-run] [--max-drafts N]
scripts\run.cmd followup [--dry-run] [--max-drafts N]
scripts\run.cmd flush        # write rows still queued from a run when Excel was open
```

Discovery segments: `roasters_adelaide`, `roasters_australia`, `freight_adelaide`, `freight_australia`. Each
week focuses on a different set of areas (e.g. Adelaide Hills, Barossa, Port Adelaide 3PLs, Melbourne,
Brisbane), so the weekly searches keep turning up new businesses.

## Development

```bash
python -m venv .venv && .venv/bin/pip install -r requirements.txt pytest
.venv/bin/python -m pytest
# IMAP integration tests against a test server (plain IMAP):
OC_TEST_IMAP=127.0.0.1:1143:user:password .venv/bin/python -m pytest tests/test_imap_integration.py
```

Code map: `discovery.py`, `leads.py`, `followup.py` (the workflows) · `workbook.py` (adaptive Excel I/O) ·
`dedupe.py` · `verify.py` (public-contact checks) · `mailer.py` (IMAP) · `compliance.py` (footer, sender checks) ·
`llm.py` + `prompts.py` (Claude) · `state.py` (SQLite) · `runtime.py` (safe commit, locks, reports).
