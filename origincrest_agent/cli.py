"""Command line: python -m origincrest_agent <command>."""

from __future__ import annotations

import argparse
import sys
import traceback

from . import __version__
from .config import ConfigError, load_config
from .llm import ClaudeAgent, log_usage
from .mailer import Mailbox
from .runtime import AgentBusy, AgentLock, ChangeSet, Report, RunContext, commit
from .state import State
from .utils import log, setup_logging, today_in
from .verify import Verifier


def _context(cfg, workflow: str, dry_run: bool) -> RunContext:
    return RunContext(
        cfg=cfg,
        state=State(cfg.data_dir / "state.db"),
        today=today_in(cfg.tz),
        report=Report(workflow, today_in(cfg.tz)),
        dry_run=dry_run,
        agent_factory=lambda: ClaudeAgent(cfg),
        mailbox_factory=lambda: Mailbox.from_config(cfg),
        verifier=Verifier(enabled=cfg["discovery"]["verify_contacts"]),
    )


def _finish(ctx: RunContext, history: bool = True) -> None:
    usage = log_usage(ctx._agent)
    path = ctx.report.write(ctx.cfg.data_dir / "reports", usage, history=history)
    log.info("Report: %s", path)


def cmd_check(cfg, args) -> int:
    from .workbook import WorkbookStore, excel_lock_reason

    ok = True
    print(f"Config: {cfg.source or '(built-in defaults - copy config.example.yaml to config.yaml)'}")
    print(f"Data folder: {cfg.data_dir}")
    for name in cfg["workbooks"]:
        path = cfg.workbook_path(name)
        try:
            store = WorkbookStore(name, path, cfg).load()
        except FileNotFoundError:
            print(f"[X] {name}: not found - {path}")
            ok = False
            continue
        lock = excel_lock_reason(path)
        print(f"[OK] {name}: {path.name}" + (f"  (currently not writable: {lock})" if lock else ""))
        for warning in store.load_warnings:
            print(f"     [!] {warning} - this Excel feature won't survive the agent's saves")
        for view in store.views():
            mapped = ", ".join(f"{k}<-'{view.headers[c]}'" for k, c in sorted(view.columns.items(), key=lambda x: x[1]))
            rows = sum(1 for _ in view.rows())
            print(f"     tab '{view.title}' [{view.kind}] header row {view.header_row}, {rows} rows")
            print(f"        columns: {mapped or '(none recognised)'}")
        if profile := cfg.profile(name):
            from .discovery import profile_categories

            notes: list[str] = []
            cats = profile_categories(cfg, store, notes)
            print(f"     new leads go to: '{profile['adelaide_tab']}' (Adelaide/SA) and '{profile['other_tab']}' "
                  "(other states)")
            for note in notes:
                print(f"     [!] {note}")
            print(f"     search categories ({len(cats)}):")
            for cat in cats:
                email = "RFQ email" if cat.kind == "freight" else "intro email"
                where = f" - e.g. {', '.join(cat.examples[:3])}" if cat.examples else ""
                print(f"        - {cat.label}  [{email}; from {cat.source}]{where}")
            ok &= bool(cats)
    from .config import expand_path

    ds = cfg["drive_sync"]
    for label, folder in (("OneDrive folder", ds["source"]), ("Google Drive folder", ds["destination"])):
        found = expand_path(folder).is_dir()
        print(("[OK] " if found else "[X] ") + f"Drive sync {label}: {folder}" + ("" if found else
              " - not found" + (" (install / start Google Drive for desktop, or fix drive_sync.destination)"
                                if label.startswith("Google") else "")))
    key_env = cfg["claude"]["api_key_env"]
    print(("[OK] " if cfg.secret(key_env) else "[X] ") + f"{key_env} " + ("set" if cfg.secret(key_env) else "missing"))
    ok &= bool(cfg.secret(key_env))
    try:
        with Mailbox.from_config(cfg) as mb:
            print(f"[OK] IMAP login {cfg['email']['username']} on {cfg['email']['imap_host']}: "
                  f"drafts='{mb.drafts}' sent={mb.sent_folders} replies={mb.reply_folders}")
            print(f"     folders on the server: {', '.join(sorted(mb.folders))}")
            if not mb.sent:
                print("[!] No Sent folder found - set email.sent_folder to one of the folders above, "
                      "otherwise follow-ups can't see what you've sent")
            if len(others := [f for f in mb.folder_candidates("drafts") if f != mb.drafts]) > 0:
                print(f"[!] More than one Drafts-like folder ({', '.join([mb.drafts] + others)}). Drafts go to "
                      f"'{mb.drafts}'; if your mail app shows another one, set email.drafts_folder to it")
            ok &= _print_sender(cfg, mb)
    except Exception as exc:  # noqa: BLE001 - report any connection/auth problem
        print(f"[X] IMAP: {exc}")
        ok = False
        _print_sender(cfg, None)
    if args.api:
        try:
            agent = ClaudeAgent(cfg)
            out = agent.structured(system="Reply with JSON.", prompt="Say ok.",
                                   schema={"type": "object", "properties": {"ok": {"type": "boolean"}},
                                           "required": ["ok"], "additionalProperties": False}, effort="low")
            print(f"[OK] Claude API ({cfg['claude']['model']}): {out}")
        except Exception as exc:  # noqa: BLE001
            print(f"[X] Claude API: {exc}")
            ok = False
    return 0 if ok else 1


def _print_sender(cfg, mailbox) -> bool:
    from .compliance import resolve_sender

    sender = resolve_sender(cfg, mailbox)
    if sender.problems:
        print("[X] Can't draft yet - set: " + ", ".join(sender.problems))
        return False
    print(f"[OK] Drafts will be from: {f'{sender.name} <{sender.email}>' if sender.name else sender.email}")
    if sender.signature:
        print(f"     signature (from {sender.signature_source}):")
        for line in sender.signature.splitlines():
            print(f"        | {line}")
    else:
        print("[!] No signature found in your recent sent emails - drafts will have none (add it when you send, "
              "or set email.signature: config)")
    return True


def cmd_drafts(cfg, args) -> int:
    """Which server folder the agent saves drafts to, and the drafts it has saved there."""
    try:
        mb = Mailbox.from_config(cfg)
        mb.connect()
    except Exception as exc:  # noqa: BLE001 - report any connection/auth problem
        print(f"[X] IMAP: {exc}")
        return 1
    try:
        how = "set by email.drafts_folder" if cfg["email"]["drafts_folder"] else "detected automatically"
        print(f"The agent saves drafts to the server folder '{mb.drafts}' ({how}).")
        folders = [mb.drafts] + [f for f in mb.folder_candidates("drafts") if f != mb.drafts]
        for folder in folders:
            saved = mb.agent_drafts(folder)
            print(f"\n'{folder}': {mb.message_count(folder)} message(s), {len(saved)} saved by the agent")
            for d in saved[:25]:
                when = d.when.astimezone().strftime("%d/%m %H:%M") if d.when else "?"
                print(f"   {when}  to {', '.join(d.to_addrs)}  \"{d.subject}\"")
        print("\nCan't see them in your mail app?")
        print(" - Outlook: drafts saved on the server are in the account's own Drafts folder (it can sit under")
        print("   Inbox), not in 'Drafts (This computer only)'. Press F9 to sync.")
        print(" - Webmail always shows the server's Drafts folder - a quick way to confirm they're there.")
        print(" - If your app uses a different folder from the list above, set email.drafts_folder to it in")
        print("   config.yaml; drafts from the next run go there.")
    finally:
        mb.close()
    return 0


def cmd_init_sheets(cfg, args) -> int:
    from .workbook import WorkbookStore

    ctx = _context(cfg, "init-sheets", args.dry_run)
    for name, profile in cfg["discovery"]["profiles"].items():
        if name not in cfg["workbooks"]:
            continue
        store = WorkbookStore(name, cfg.workbook_path(name), cfg).load()
        for tab in (profile["adelaide_tab"], profile["other_tab"]):
            store.get_or_create(tab, store.default_kind)
        if args.dry_run:
            print(f"{store.path.name}: would create missing tabs")
            continue
        store.save(ctx.backup_dir, cfg["excel"]["backup_keep"])
        print(f"{store.path.name}: tabs ready")
    return 0


def run_workflow(cfg, args) -> int:
    from .discovery import run_discover
    from .followup import run_followup
    from .leads import run_leads

    ctx = _context(cfg, args.command, args.dry_run)
    run_id = ctx.state.start_run(args.command)
    status = 0
    # drive-sync runs every few minutes: if the agent is busy it just tries again at the next check.
    quick_check = args.command == "drive-sync"
    busy = False
    try:
        with AgentLock(cfg.data_dir / "agent.lock", wait_minutes=0 if quick_check else 45):
            if args.command == "discover":
                split = lambda v: [x.strip() for x in v.split(",") if x.strip()] if v else None  # noqa: E731
                run_discover(ctx, workbooks=split(args.workbook), categories=split(args.categories),
                             regions=split(args.region), max_rows=args.max_rows)
                if cfg["discovery"]["run_leads_after"] and not args.no_leads:
                    run_leads(ctx)
            elif args.command == "leads":
                run_leads(ctx, max_drafts=args.max_drafts)
            elif args.command == "followup":
                run_followup(ctx, max_drafts=args.max_drafts)
            elif args.command == "drive-sync":
                from .drive_sync import run_drive_sync

                run_drive_sync(ctx)
            elif args.command == "flush":
                commit(ctx, ChangeSet())
    except AgentBusy as exc:
        busy = True
        ctx.report.add("Skipped", f"Not run: {exc}")
    except Exception as exc:  # noqa: BLE001 - log every failure to the report
        ctx.report.add("Errors", f"{type(exc).__name__}: {exc}")
        log.debug(traceback.format_exc())
        status = 1
    finally:
        nothing_happened = busy or set(ctx.report.sections) == {"Drive sync"}
        if quick_check and nothing_happened and status == 0:
            ctx.state.discard_run(run_id)  # no changes: no report, no log noise
        else:
            # A failing quick check (e.g. Google Drive not running) updates latest_drive-sync.md only.
            _finish(ctx, history=not (quick_check and status))
            ctx.state.finish_run(run_id, ctx.report.counts)
        ctx.state.close()
    return status


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="origincrest_agent", description="OriginCrest outreach agent")
    parser.add_argument("--config", help="path to config.yaml")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("check", help="check config, workbooks, IMAP login and API key")
    p.add_argument("--api", action="store_true", help="also make a tiny Claude API call")
    sub.add_parser("drafts", help="show which mail folder drafts are saved to, and the drafts saved there")
    p = sub.add_parser("init-sheets", help="create the new-lead tabs (Adelaide / other states) in each workbook")
    p.add_argument("--dry-run", action="store_true")
    p = sub.add_parser("discover", help="workflow 1: find new leads for each workbook")
    p.add_argument("--dry-run", action="store_true", help="search and report, but write nothing")
    p.add_argument("--workbook", help="only this workbook, e.g. rfq or master")
    p.add_argument("--categories", help="only these categories (comma list, as shown by check)")
    p.add_argument("--region", choices=["adelaide", "other"], help="only Adelaide/SA or only the other states")
    p.add_argument("--max-rows", type=int, help="cap on new rows per workbook for this run")
    p.add_argument("--no-leads", action="store_true", help="don't run workflow 2 afterwards")
    p = sub.add_parser("leads", help="workflow 2: research new leads and draft intro / RFQ emails")
    p.add_argument("--dry-run", action="store_true", help="research and preview drafts as .eml files only")
    p.add_argument("--max-drafts", type=int)
    p = sub.add_parser("followup", help="workflow 3: check replies and draft follow-ups")
    p.add_argument("--dry-run", action="store_true", help="check and preview drafts as .eml files only")
    p.add_argument("--max-drafts", type=int)
    p = sub.add_parser("drive-sync", help="copy new/changed OneDrive Origincrest files to Google Drive")
    p.add_argument("--dry-run", action="store_true", help="list what would be copied, copy nothing")
    p = sub.add_parser("flush", help="write any rows still queued (e.g. after Excel was open)")
    p.add_argument("--dry-run", action="store_true")

    args = parser.parse_args(argv)
    try:
        cfg = load_config(args.config)
    except ConfigError as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        return 2
    setup_logging(cfg.data_dir / "logs", verbose=args.verbose)
    if args.command == "check":
        return cmd_check(cfg, args)
    if args.command == "init-sheets":
        return cmd_init_sheets(cfg, args)
    if args.command == "drafts":
        return cmd_drafts(cfg, args)
    return run_workflow(cfg, args)
