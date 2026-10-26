# OriginCrest outreach agent

One agent with three workflows. It runs on your Windows PC, reads and writes the two
OneDrive workbooks directly, uses Claude (with web search) for research and writing, and
saves every email as an **IMAP draft** for you to review. It never sends email.

| # | Workflow | When | What it does |
|---|----------|------|--------------|
| 1 | `discover` | Weekly (Mon 6:30) | Finds new leads on the web, separately for each workbook (below). Skips anything already in that workbook, then adds new rows with status **New - to review** to two tabs: one for Adelaide / SA, one for all other states. |
| 2 | `leads` | Every 2 h, 7:00-19:00 (and right after discovery) | Turns each new lead into a draft written from its own row. Freight forwarders, customs brokers and 3PLs get a tailored **RFQ**. Coffee businesses are researched first (what they sell, size, sourcing, current Indonesian coffees, fit 0-100 for Sumatra / Sulawesi / Java / Flores), then get a personalised **intro email**. |
| 3 | `followup` | Weekdays 7:30 | Reads the RFQ workbook and every new-lead tab, finds **Open** rows with no reply after 5 days (quotes) or 10 days (leads), checks your mailbox for replies first, then drafts a threaded follow-up and updates **Last contacted** and **Next action date**. |

| Workbook | New leads it gets | Tabs | Email |
|----------|-------------------|------|-------|
| `OriginCrest_Freight_3PL_Contacts_RFQ.xlsx` | Freight forwarders, customs brokers and 3PLs that handle green or roasted coffee | `New Freight-3PL - Adelaide`, `New Freight-3PL - Other States` | RFQ |
| `Australian_Coffee_3PL_Master_Contacts_Oct2026.xlsx` | Every category in the **New Prospects Oct-26** tab, plus the importer / roaster / wholesaler / trader categories in **All contacts** (read from their Category column each run) | `New Leads - Adelaide`, `New Leads - Other States` | Intro (RFQ for any logistics category) |

`scripts\run.cmd check` lists the categories it found in your tabs, with an example company for each.

The spreadsheets stay where they are, so your existing OneDrive → Google Drive backup keeps working.

## Guardrails

- **Drafts only.** Emails are saved to your Drafts folder. Nothing goes out until you press Send.
- **Discovery rows are quarantined.** New rows get status `New - to review`, `Added by = OriginCrest Agent`
  and a `Source URL`, on their own "New ..." tabs. Only rows in those tabs are ever drafted to: your
  existing tabs (`All contacts`, `New Prospects Oct-26`, your RFQ tracker) are read for categories,
  duplicates and RFQ examples, never emailed.
- **Only publicly listed business contacts.** Claude is told never to guess or build email addresses.
  The agent then loads the source page (and the company's contact page) and **keeps an email, phone or
  contact name only if it actually appears there**. If the page says "no unsolicited email", the email
  is left blank. Anything dropped is explained in `Agent notes`.
- **Spam Act 2003.** Every email carries your own signature, copied from your recent sent emails, so
  recipients see who you are and how to reach you. It also gets an opt-out, added by code rather than by the
  model: intro emails get the full unsubscribe footer, RFQs a one-line "reply 'unsubscribe' and we won't
  contact you again". Both carry a `List-Unsubscribe` header. Unsubscribe replies are detected on the next weekday run, the row is set to
  `Unsubscribed - do not contact`, and the address goes on a suppression list that every workflow checks.
  If no From address can be found, the agent won't draft anything until one is set in `config.yaml`.
- **Caps per run:** 20 new rows per workbook (6 per category and region), 4 categories per week (rotating),
  10 companies researched, 8 drafts, 15 follow-up drafts, max 3 follow-ups per contact. All adjustable.
- **Your data is safe:** a timestamped backup of each workbook is made before every save, nothing is written
  while a workbook is open in Excel, and your own columns (e.g. `Notes`) are never overwritten. The agent
  writes its comments to its own `Agent notes` column.
- **Leaves your contacts alone:** within the new-lead tabs, workflow 2 only touches rows whose status is blank,
  `New`, `New - to review`, `Reviewed`, `Approved`, `To contact`, `Lead` or `Prospect`. Before drafting it
  also checks your Sent folder: if you've already emailed them, the row is set to `Open` instead.
- **No invented facts.** Emails use only the contact's row, the research on them and the OriginCrest details
  you allow. RFQ shipment details come from `rfq.requirements`, your RFQ tracker rows and the last RFQ you
  sent; anything still unknown is left as a `[placeholder]` to fill in, and the row's notes say so.

> Not legal advice: the Spam Act lets you email a business address that is conspicuously published,
> if the message relates to their business and the page doesn't say they don't want it. The Act also
> restricts using "address-harvesting" software to build mailing lists. That's why this agent keeps a
> person in the loop: you review every row and every draft, it researches companies one at a time, it records
> where each address was published, and it sends nothing in bulk. See ACMA's Spam Act guidance if unsure.

## How the statuses move

```
discover ─► New - to review ─► (leads) ─┬─► Intro drafted - to review ─(you send it)─► Open
                                        ├─► RFQ drafted - to review ───(you send it)─► Open
                                        ├─► Researched - low fit        (set Approved to draft anyway)
                                        └─► No public email

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
- Quote vs lead: taken from the `Record type` column (the agent fills it in for new leads), otherwise from the
  Category (logistics words like freight / 3PL / customs mean quote), otherwise from the workbook: RFQ workbook
  rows are quotes, master workbook rows are leads.
- **Talking points:** the new-lead tabs have a `Talking points` column. Anything you write there (how you met,
  what to mention) is built into that contact's email.

## Columns

Existing columns are matched by name (e.g. `Company Name`, `Contact Person`, `Email Address`, `Date Contacted`),
and the header row doesn't have to be row 1. Missing columns are added at the right-hand end of the header row,
styled like your existing headers, and Excel tables and filters are stretched to include them.

| Workflow | Columns it may add |
|----------|--------------------|
| 1 | Company, Category, Contact name, Role, Email, Phone, Website, City, State, Services / offering, Talking points, Source URL, Status, Date added, Added by, Record type, Agent notes |
| 2 | Draft created, Draft subject; for coffee businesses also Coffee offering, Size estimate, Sourcing, Current Indonesian coffees, Fit score (0-100), Best-fit origins, Fit rationale, Researched on |
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
5. **Email (IMAP)**: put the mailbox password in `.env` as `IMAP_PASSWORD=...`. Use an app password if
   your mail host provides them.
6. **Edit `config.yaml`**:
   - `email`: your mail server's IMAP settings, the same ones your phone or Outlook uses for this mailbox:
     `imap_host` (e.g. `mail.yourdomain.com.au`), `imap_port` + `security` (`993` + `ssl`, or `143` +
     `starttls`), and `username` (usually the full email address).
   - Signature: nothing to do. The agent copies the signature you already use from your recent sent emails,
     with your name and From address, and `check` shows what it found. Text signatures are copied exactly;
     a logo image in an HTML signature isn't, because drafts are plain text. If you'd rather paste your
     own each time, set `email.signature: none`.
   - `rfq.requirements` (optional): what you need quoted, e.g. lanes, container size, volumes, services. Leave
     it blank and the agent works it out from your RFQ tracker and last RFQ email.
   - `offer` (optional): extra facts about OriginCrest the emails may use, such as sample sizes or where stock
     is held. Left blank, emails describe OriginCrest as a supplier of Indonesian single-origin green coffee.
   - The workbook paths are already set to your two files.
7. **Check everything**:
   ```powershell
   scripts\run.cmd check --api
   ```
   This lists every tab, the header row it found and how each column was matched, and the categories it will
   search for each workbook, with an example company from your sheet. It also logs in to your
   mailbox and shows which folders it will use for Drafts, Sent and replies, plus every folder on the server.
   Fix anything marked `[X]` or `[!]`.
8. **First run by hand**, with Excel closed:
   ```powershell
   scripts\run.cmd discover --workbook rfq --region adelaide --max-rows 3 --dry-run   # writes nothing
   scripts\run.cmd discover --workbook rfq --region adelaide --max-rows 3            # 3 rows + RFQ drafts
   scripts\run.cmd discover --workbook master --categories Roasters --max-rows 3    # 3 rows + intros
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

- **Can't see the drafts?** Run `scripts\run.cmd drafts`. It shows the server folder the agent saves to and
  lists the drafts in it. In Outlook, server drafts are in the account's own Drafts folder (sometimes under
  Inbox), not "Drafts (This computer only)". Webmail always shows them. If your mail app uses a different
  folder, set `email.drafts_folder` to it.

- **Email setup.** The agent works with any standard IMAP server. Drafts appear in every app connected to the
  mailbox by IMAP (Outlook, Thunderbird, phone, webmail). Two things matter for follow-ups:
  - your mail apps must use **IMAP, not POP3**, or replies get downloaded off the server before the agent
    sees them;
  - sent mail must be saved to the server's Sent folder (IMAP apps do this by default).

  If replies are filed by mail rules, list those folders in `email.reply_folders`. A certificate error usually
  means the host name doesn't match the server's certificate: on shared hosting, use the server name your
  host gives for mail apps (e.g. `server123.hostname.com.au`).

- **Excel open?** The agent never writes to a workbook that's open. It waits up to 5 minutes, then keeps
  the changes (new rows are queued; research and drafts are remembered) and writes them on the next run.
  Close the workbooks overnight and the early-morning runs will always get through.
- **No double drafts.** Research, drafts and follow-ups are recorded in a local database
  (`%LOCALAPPDATA%\OriginCrestAgent\state.db`), so re-runs never draft the same email twice. A follow-up isn't
  drafted while an earlier draft to the same address is still waiting unsent in Drafts.
- **Deleted a junk discovery row?** It won't be re-added: the agent remembers what it found before.
- **Same company in both workbooks** is drafted to only once.
- **Workflow 3** reads every tab of the RFQ workbook, plus the new-lead tabs of the master workbook so the
  leads the agent found there get followed up. Add `master` to `followup.workbooks` if you want every Open row
  in the master workbook chased too.
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
scripts\run.cmd discover [--dry-run] [--workbook rfq|master] [--categories "Roasters,Importers"] [--region adelaide|other] [--max-rows N] [--no-leads]
scripts\run.cmd leads [--dry-run] [--max-drafts N]
scripts\run.cmd followup [--dry-run] [--max-drafts N]
scripts\run.cmd flush        # write rows still queued from a run when Excel was open
scripts\run.cmd drafts       # which mail folder drafts go to, and the drafts the agent saved there
```

Each search is one category in one region (Adelaide/SA or other states). Each week focuses on different areas
(e.g. Adelaide Hills, Barossa, Melbourne, Brisbane) or, for freight, different services (Port Adelaide
forwarders, customs brokers, food-grade 3PLs), so the weekly searches keep turning up new businesses.

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
