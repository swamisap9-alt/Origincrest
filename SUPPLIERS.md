# OriginCrest supplier agent

Finds Indonesian green-coffee suppliers (exporters, cooperatives with an export licence, mills),
compares each with your current vendor, and saves a sourcing enquiry for each one in your Drafts folder.
It never sends email.

It is a **separate agent** from the outreach agent in [README.md](README.md):

| | Outreach agent | Supplier agent |
|---|---|---|
| Settings | `config.yaml` | `config.suppliers.yaml` |
| Run by hand | `scripts\run.cmd ...` | `scripts\suppliers.cmd ...` |
| Scheduled tasks | Task Scheduler folder `OriginCrest` | Task Scheduler folder `OriginCrest Suppliers` |
| Install / remove tasks | `install_tasks.ps1` / `uninstall_tasks.ps1` | `install_supplier_tasks.ps1` / `uninstall_supplier_tasks.ps1` |
| Workbooks | the RFQ and master workbooks | `Indonesian_Coffee_Suppliers.xlsx` only |
| Data folder (state, logs, reports, backups, run lock) | `%LOCALAPPDATA%\OriginCrestAgent` | `%LOCALAPPDATA%\OriginCrestAgent\suppliers` |

Editing one agent's settings can't change or break the other, and neither touches the other's workbooks,
tasks or history. They run independently (each has its own run lock).

**Shared, read only:** the supplier agent uses the same mailbox as the outreach agent, so drafts land in the
same Drafts folder with the same From address and signature. It reads the `email`, `sender`, `claude` and
`timezone` sections of `config.yaml`, and the keys in `.env`. It never writes `config.yaml` (nothing in the
agent writes settings files). To use different values for the supplier agent, put that section in
`config.suppliers.yaml`; it wins.

> **This repository is public.** Your vendor, customer, demand and prices go in `config.suppliers.yaml`,
> which is never committed. Don't put them in `config.suppliers.example.yaml` or elsewhere in this folder.

| Workflow | Command | Scheduled | What it does |
|---|---|---|---|
| Search | `search` | every 3 days, 6:00 | Searches two regions or supplier types per run (rotating; with `priority_focus` set, the next priority search plus one from the rotation). Adds new suppliers to `New Suppliers - Indonesia` with contact details that appear on the supplier's own pages, origins, certifications, export capability, a fit score, **Comparison with current vendor** and **Suggested approach**. Then drafts an enquiry to each. |
| Follow-up | `followup` | weekdays, 7:45 | Reads the supplier workbook, picks up enquiries you've sent (from your Sent folder), checks for replies, and drafts a threaded follow-up when there's no reply after `followup.supplier_days` (3). Up to 3 per supplier. |
| Draft | `draft` / `redraft "Name"` | after each search | Enquiries for suppliers you added to the tab yourself, or a new first draft after you deleted one. |

## Setup (about 5 minutes; the outreach agent must already be set up)

1. Get the latest code (`git pull`), then create the settings file:
   ```powershell
   Copy-Item config.suppliers.example.yaml config.suppliers.yaml
   ```
2. Fill in `config.suppliers.yaml`:
   - `suppliers.baseline_vendor`: your current vendor's name and what you know about them (`profile`).
     Every new supplier is compared with it. It is never put in an email.
   - `sourcing.demand`: the order you need supplied (the trial by LCL, then the FCL step), word for word as it
     should appear in the enquiries. Left blank, drafts carry `[placeholders]`.
   - `suppliers.priority_focus` (optional): searches for what you need now, e.g. the coffees in a customer's
     quote. One is searched every run until you remove them.
3. Check, then try it with Excel closed:
   ```powershell
   scripts\suppliers.cmd check
   scripts\suppliers.cmd search --dry-run --max-rows 3    # searches and reports, writes nothing
   scripts\suppliers.cmd search --max-rows 3              # 3 suppliers + enquiry drafts
   scripts\suppliers.cmd followup --dry-run               # previews follow-ups as .eml files
   ```
4. Schedule it:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\install_supplier_tasks.ps1
   ```
   Change times with `-SearchEveryDays 3 -SearchTime 06:00 -FollowupTime 07:45`. This only adds or replaces
   the two supplier tasks; your outreach tasks and their times are left alone. (It also removes
   `OriginCrest\5 Suppliers`, which an earlier version added to the outreach folder.)

## How the statuses move

```
search ─► New - to review ─► (draft) ─┬─► Enquiry drafted - to review ─(you send it)─► Open
                                      ├─► Researched - low fit   (fit under 50; set Approved to draft anyway)
                                      └─► No public email        (add one and set Approved)

Open ─(followup, no reply after 3 days)─► follow-up draft, Last contacted = today, Next action date = +3
Open ─(reply found)─────────────────────► Replied - review
```

- **Covers current demand** (Yes / Partly / No) says whether the supplier sells the coffees in
  `sourcing.demand`. Yes or Partly: the enquiry describes that order. No: the enquiry introduces OriginCrest and
  asks about the supplier's own coffees instead, without naming your customer.
- **Searches** (top of the run report) lists each search with how many suppliers it found and why candidates
  were left out, so an empty search is never silent. Searches run in Indonesian as well as English, and the
  Ministry of Trade's exporter directory (InaExport) counts as a source.
- **Private columns:** `Comparison with current vendor`, `Suggested approach` and `Fit rationale` are for you.
  The email writer never sees them, and enquiries never mention another supplier or your vendor.
- **Contact details** are kept only if they appear on the supplier's own pages; anything dropped is explained
  in `Agent notes`. WhatsApp numbers go in the `Phone` column.
- **No unsubscribe footer:** supplier enquiries are purchase enquiries, not marketing, so they carry your
  signature only. A supplier who replies "unsubscribe" is still never contacted again.
- **Duplicates:** a supplier is added once (`PT` / `CV` / `UD` and `Tbk` are ignored when comparing names), your
  vendor is never added, and deleting a row doesn't bring it back.
- Reports: `%LOCALAPPDATA%\OriginCrestAgent\suppliers\reports\latest_supplier-search.md` (and
  `latest_supplier-followup.md`). Log: `...\suppliers\logs\agent.log`.

## Commands

```
scripts\suppliers.cmd check [--api]
scripts\suppliers.cmd search [--dry-run] [--max-rows N] [--max-drafts N] [--focus "..."] [--no-drafts]
scripts\suppliers.cmd draft [--dry-run] [--max-drafts N]
scripts\suppliers.cmd redraft "Kopi Gayo" [--dry-run]
scripts\suppliers.cmd followup [--dry-run] [--max-drafts N]
scripts\suppliers.cmd flush      # write rows still queued from a run when Excel was open
```

Code: `supplier_agent.py` (commands), `supplier_config.py` (settings), `suppliers.py` (search, drafts,
follow-up), `supplier_prompts.py`. Drafting and follow-ups reuse the outreach agent's IMAP code
(`mailer.py`, `leads.py`, `followup.py`), so drafts behave exactly like intros and RFQs.
