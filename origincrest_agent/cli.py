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
    from .compliance import sender_problems

    if problems := sender_problems(cfg):
        print("[!] Fill in before drafting emails: " + ", ".join(problems))
    else:
        print("[OK] Sender details complete")
    key_env = cfg["claude"]["api_key_env"]
    print(("[OK] " if cfg.secret(key_env) else "[X] ") + f"{key_env} " + ("set" if cfg.secret(key_env) else "missing"))
    ok &= bool(cfg.secret(key_env))
    try:
        with Mailbox.from_config(cfg) as mb:
            print(f"[OK] IMAP login {cfg['email']['username']}: drafts='{mb.drafts}' sent='{mb.sent}' "
                  f"replies={mb.reply_folders}")
    except Exception as exc:  # noqa: BLE001 - report any connection/auth problem
        print(f"[X] IMAP: {exc}")
        ok = False
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


def cmd_init_sheets(cfg, args) -> int:
    from .workbook import WorkbookStore

    ctx = _context(cfg, "init-sheets", args.dry_run)
    tabs = cfg["tabs"]
    for name in cfg["discovery"]["workbooks"]:
        store = WorkbookStore(name, cfg.workbook_path(name), cfg).load()
        for key, kind in (("roasters_australia", "roaster"), ("roasters_adelaide", "roaster"),
                          ("freight_australia", "freight"), ("freight_adelaide", "freight")):
            store.get_or_create(tabs[key], kind)
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
                segments = args.segments.split(",") if args.segments else None
                run_discover(ctx, segments=segments, max_rows=args.max_rows)
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
    p = sub.add_parser("init-sheets", help="create the discovery tabs in both workbooks")
    p.add_argument("--dry-run", action="store_true")
    p = sub.add_parser("discover", help="workflow 1: find new roasters and forwarders/3PLs")
    p.add_argument("--dry-run", action="store_true", help="search and report, but write nothing")
    p.add_argument("--max-rows", type=int, help="override the per-run cap on new rows")
    p.add_argument("--segments", help="comma list, e.g. roasters_adelaide,freight_adelaide")
    p.add_argument("--no-leads", action="store_true", help="don't run workflow 2 afterwards")
    p = sub.add_parser("leads", help="workflow 2: research new roaster rows and draft intros")
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
