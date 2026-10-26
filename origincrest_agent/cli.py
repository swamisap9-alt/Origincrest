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


def _finish(ctx: RunContext) -> None:
    usage = log_usage(ctx._agent)
    path = ctx.report.write(ctx.cfg.data_dir / "reports", usage)
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
    key_env = cfg["claude"]["api_key_env"]
    print(("[OK] " if cfg.secret(key_env) else "[X] ") + f"{key_env} " + ("set" if cfg.secret(key_env) else "missing"))
    ok &= bool(cfg.secret(key_env))
    try:
        with Mailbox.from_config(cfg) as mb:
            print(f"[OK] IMAP login {cfg['email']['username']} on {cfg['email']['imap_host']}: "
                  f"drafts='{mb.drafts}' sent='{mb.sent}' replies={mb.reply_folders}")
            print(f"     folders on the server: {', '.join(sorted(mb.folders))}")
            if not mb.sent:
                print("[!] No Sent folder found - set email.sent_folder to one of the folders above, "
                      "otherwise follow-ups can't see what you've sent")
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
    print(f"[OK] Drafts will be from: {sender.name} <{sender.email}>")
    if sender.signature:
        print(f"     signature (from {sender.signature_source}):")
        for line in sender.signature.splitlines():
            print(f"        | {line}")
    else:
        print("[!] No signature found in your Sent folder - drafts will have none (add it when you send, "
              "or set email.signature: config)")
    return True


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
    try:
        with AgentLock(cfg.data_dir / "agent.lock"):
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
            elif args.command == "flush":
                commit(ctx, ChangeSet())
    except AgentBusy as exc:
        ctx.report.add("Skipped", f"Not run: {exc}")
    except Exception as exc:  # noqa: BLE001 - log every failure to the report
        ctx.report.add("Errors", f"{type(exc).__name__}: {exc}")
        log.debug(traceback.format_exc())
        status = 1
    finally:
        _finish(ctx)
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
    return run_workflow(cfg, args)
