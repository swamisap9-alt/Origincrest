"""Command line for the supplier agent: python -m origincrest_agent.supplier_agent <command>.

Separate from the outreach agent (python -m origincrest_agent): its own settings file
(config.suppliers.yaml), data folder, run lock, reports and scheduled tasks. It shares the mailbox:
email / sender / claude / timezone are read from config.yaml, which it never writes.
"""

from __future__ import annotations

import argparse
import sys
import traceback

from datetime import datetime

from .cli import _context, _finish, _print_sender
from .config import DEFAULTS, ConfigError, expand_path
from .llm import ClaudeAgent
from .mailer import Mailbox
from .prompts import _filled
from .runtime import AgentBusy, AgentLock, ChangeSet, commit
from .supplier_config import SUPPLIER_CONFIG_NAME, WORKBOOK, load_supplier_config, supplier_config_path
from .utils import keep_awake, log, setup_logging, today_in

COMMANDS_NEEDING_SETTINGS = ("search", "draft", "redraft", "followup")


def _config_failure(message: str) -> None:
    """Show a settings problem and log it in the supplier agent's own log (scheduled runs have no window)."""
    print(message, file=sys.stderr)
    try:
        log_dir = expand_path(DEFAULTS["data_dir"]) / "suppliers" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        with open(log_dir / "agent.log", "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} ERROR   {message}\n")
    except OSError:
        pass


def cmd_check(cfg, args) -> int:
    from .suppliers import run_focus
    from .workbook import WorkbookStore, excel_lock_reason

    ok = True
    if cfg.source:
        print(f"[OK] Supplier agent settings: {cfg.source}")
    else:
        print(f"[X] Supplier agent settings: {supplier_config_path(args.config)} not found - create it with")
        print(f"       Copy-Item config.suppliers.example.yaml {SUPPLIER_CONFIG_NAME}")
        ok = False
    for key in cfg.unknown_keys:
        print(f"     [!] '{key}:' in {SUPPLIER_CONFIG_NAME} isn't a supplier agent setting and is ignored")
    if cfg.shared_source:
        used = ", ".join(cfg.shared_used) or "nothing (all set in the supplier file)"
        print(f"[OK] Shared mailbox settings read from {cfg.shared_source} (read only): {used}")
    else:
        print("[!] config.yaml not found - set email / sender / claude in the supplier file instead")
    for section in cfg.ignored_in_shared:
        print(f"     [!] config.yaml has a '{section}:' section. The supplier agent ignores it (and so does the "
              f"outreach agent): move it to {SUPPLIER_CONFIG_NAME}, then delete it from config.yaml")
    for warning in cfg.warnings:
        print(f"     [!] {warning}")
    print(f"Data folder (state, logs, reports, backups): {cfg.data_dir}")

    sc = cfg["suppliers"]
    path = cfg.workbook_path(WORKBOOK)
    if path.exists():
        store = WorkbookStore(WORKBOOK, path, cfg).load()
        lock = excel_lock_reason(path)
        print(f"[OK] Supplier workbook: {path}" + (f"  (currently not writable: {lock})" if lock else ""))
        for view in store.views():
            print(f"     tab '{view.title}': {sum(1 for _ in view.rows())} suppliers")
    else:
        print(f"[OK] Supplier workbook will be created by the first search: {path}")
    print(f"Search: every {sc['every_days']} days, tab '{sc['tab']}'; follow-ups after "
          f"{cfg['followup']['supplier_days']} days (up to {cfg['followup']['max_followups_per_contact']})")
    for i, theme in enumerate(run_focus(cfg, today_in(cfg.tz)), start=1):
        print(f"     next search {i}: {theme}")
    vendor = sc["baseline_vendor"]
    for label, value, hint in (
            (f"compared with {_filled(vendor.get('name')) or '(no vendor named)'}", vendor.get("profile"),
             "suppliers.baseline_vendor.profile is blank - comparisons will be one-sided"),
            ("demand in the enquiries", cfg["sourcing"].get("demand"),
             "sourcing.demand is blank - drafts will carry [placeholders] for the LCL / FCL quantities")):
        text = _filled(value).replace("\n", " ")
        print(f"[OK] {label}: {text[:100]}{'...' if len(text) > 100 else ''}" if text else f"[!] {hint}")

    key_env = cfg["claude"]["api_key_env"]
    print(("[OK] " if cfg.secret(key_env) else "[X] ") + f"{key_env} " + ("set" if cfg.secret(key_env) else "missing"))
    ok &= bool(cfg.secret(key_env))
    try:
        with Mailbox.from_config(cfg) as mb:
            print(f"[OK] IMAP login {cfg['email']['username']} on {cfg['email']['imap_host']}: "
                  f"drafts='{mb.drafts}' sent={mb.sent_folders} replies={mb.reply_folders}")
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


def run_workflow(cfg, args) -> int:
    from .suppliers import run_supplier_drafts, run_supplier_followup, run_suppliers

    workflow = f"supplier-{args.command}"
    ctx = _context(cfg, workflow, args.dry_run)
    run_id = ctx.state.start_run(workflow)
    status = 0
    try:
        with AgentLock(cfg.data_dir / "agent.lock", wait_minutes=45), keep_awake():
            if args.command == "search":
                run_suppliers(ctx, max_rows=args.max_rows, focus=[args.focus] if args.focus else None)
                if not args.no_drafts:
                    run_supplier_drafts(ctx, max_drafts=args.max_drafts)
            elif args.command == "draft":
                run_supplier_drafts(ctx, max_drafts=args.max_drafts)
            elif args.command == "redraft":
                run_supplier_drafts(ctx, redraft=args.companies)
            elif args.command == "followup":
                run_supplier_followup(ctx, max_drafts=args.max_drafts)
            elif args.command == "flush":
                commit(ctx, ChangeSet())
    except AgentBusy as exc:
        ctx.report.add("Skipped", f"Not run: {exc}")
    except Exception as exc:  # noqa: BLE001 - log every failure to the report
        ctx.report.add("Errors", f"{type(exc).__name__}: {exc}")
        log.debug(traceback.format_exc())
        status = 1
    finally:
        for section in cfg.ignored_in_shared:
            ctx.report.add("Settings", f"config.yaml has a '{section}:' section, which is ignored: move it to "
                                       f"{SUPPLIER_CONFIG_NAME}")
        _finish(ctx)
        ctx.state.finish_run(run_id, ctx.report.counts)
        ctx.state.close()
    return status


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="origincrest_agent.supplier_agent",
                                     description="OriginCrest supplier agent: Indonesian green-coffee suppliers")
    parser.add_argument("--config", help=f"path to {SUPPLIER_CONFIG_NAME}")
    parser.add_argument("--shared-config", help="path to the config.yaml whose email / sender / claude settings "
                                                "are used (read only)")
    parser.add_argument("-v", "--verbose", action="store_true")
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("check", help="check settings, workbook, IMAP login and API key")
    p.add_argument("--api", action="store_true", help="also make a tiny Claude API call")
    p = sub.add_parser("search", help="find suppliers, compare them with your vendor, then draft enquiries")
    p.add_argument("--dry-run", action="store_true", help="search and report, but write nothing")
    p.add_argument("--max-rows", type=int, help="cap on new suppliers for this run")
    p.add_argument("--max-drafts", type=int)
    p.add_argument("--focus", help='search this instead of the rotation, e.g. "Toraja exporters in Makassar"')
    p.add_argument("--no-drafts", action="store_true", help="don't draft enquiries afterwards")
    p = sub.add_parser("draft", help="draft enquiries for suppliers in the workbook that don't have one yet")
    p.add_argument("--dry-run", action="store_true", help="preview the drafts as .eml files only")
    p.add_argument("--max-drafts", type=int)
    p = sub.add_parser("redraft", help="write a new first enquiry for suppliers whose draft you deleted")
    p.add_argument("companies", nargs="+", help='company names as in the sheet, e.g. "Kopi Gayo"')
    p.add_argument("--dry-run", action="store_true")
    p = sub.add_parser("followup", help="check replies and draft follow-ups to unanswered enquiries")
    p.add_argument("--dry-run", action="store_true", help="check and preview drafts as .eml files only")
    p.add_argument("--max-drafts", type=int)
    p = sub.add_parser("flush", help="write any supplier rows still queued (e.g. after Excel was open)")
    p.add_argument("--dry-run", action="store_true")

    args = parser.parse_args(argv)
    try:
        cfg = load_supplier_config(args.config, args.shared_config)
    except ConfigError as exc:
        _config_failure(f"Supplier agent config error: {exc}")
        return 2
    setup_logging(cfg.data_dir / "logs", verbose=args.verbose)
    try:
        if args.command == "check":
            return cmd_check(cfg, args)
        if args.command in COMMANDS_NEEDING_SETTINGS and cfg.source is None:
            message = (f"{supplier_config_path(args.config)} not found. Create it from the example "
                       f"(Copy-Item config.suppliers.example.yaml {SUPPLIER_CONFIG_NAME}) and fill it in.")
            print(message, file=sys.stderr)
            log.error(message)
            return 2
        return run_workflow(cfg, args)
    except Exception:  # noqa: BLE001 - scheduled runs have no window to show this in, so log it
        log.exception("supplier agent '%s' stopped with an unexpected error", args.command)
        return 1


if __name__ == "__main__":
    sys.exit(main())
