import pytest

from origincrest_agent.drive_sync import run_drive_sync


@pytest.fixture
def folders(cfg, tmp_path):
    src, dst = tmp_path / "OneDrive Origincrest", tmp_path / "My Drive" / "Origincrest files synced from laptop"
    (src / "OriginCrest enquiries").mkdir(parents=True)
    dst.mkdir(parents=True)
    (src / "notes.txt").write_text("root file")
    (src / "OriginCrest enquiries" / "Master.xlsx").write_bytes(b"xlsx v1")
    (src / "OriginCrest enquiries" / "~$Master.xlsx").write_text("excel lock")
    (src / "desktop.ini").write_text("x")
    (src / ".drive-sync-manifest.json").write_text("{}")  # the old cloud job's manifest
    cfg.data["drive_sync"].update(source=str(src), destination=str(dst))
    return src, dst


def test_copies_new_files_and_skips_lock_files(folders, make_ctx):
    src, dst = folders
    ctx = make_ctx("drive-sync")
    assert run_drive_sync(ctx) == 2
    assert (dst / "notes.txt").read_text() == "root file"
    assert (dst / "OriginCrest enquiries" / "Master.xlsx").read_bytes() == b"xlsx v1"
    assert not (dst / "OriginCrest enquiries" / "~$Master.xlsx").exists()
    assert not (dst / "desktop.ini").exists() and not (dst / ".drive-sync-manifest.json").exists()
    assert ctx.report.sections["New Drive folders"] == ["OriginCrest enquiries"]

    ctx = make_ctx("drive-sync")
    assert run_drive_sync(ctx) == 0
    assert ctx.report.sections["Drive sync"] == ["No changes"]


def test_changed_file_is_replaced_and_nothing_is_deleted(folders, make_ctx):
    src, dst = folders
    run_drive_sync(make_ctx("drive-sync"))
    (src / "OriginCrest enquiries" / "Master.xlsx").write_bytes(b"xlsx v2 - edited")
    (src / "notes.txt").unlink()
    ctx = make_ctx("drive-sync")
    assert run_drive_sync(ctx) == 1
    assert (dst / "OriginCrest enquiries" / "Master.xlsx").read_bytes() == b"xlsx v2 - edited"
    assert ctx.report.sections["Updated in Drive (new version)"] == ["OriginCrest enquiries/Master.xlsx"]
    assert (dst / "notes.txt").exists()  # deleted locally, kept in Drive
    assert ctx.report.sections["Missing locally (kept in Drive)"] == ["notes.txt"]
    ctx2 = make_ctx("drive-sync")
    run_drive_sync(ctx2)
    assert "Missing locally (kept in Drive)" not in ctx2.report.sections  # reported once


def test_files_already_in_drive_are_adopted_not_recopied(folders, make_ctx):
    src, dst = folders
    (dst / "notes.txt").write_text("root file")  # uploaded earlier by the cloud sync
    ctx = make_ctx("drive-sync")
    assert run_drive_sync(ctx) == 1  # only Master.xlsx
    assert ctx.report.sections["Added to Drive"] == ["OriginCrest enquiries/Master.xlsx"]


def test_dry_run_and_missing_drive_folder(folders, make_ctx, cfg, tmp_path):
    src, dst = folders
    ctx = make_ctx("drive-sync", dry_run=True)
    assert run_drive_sync(ctx) == 0
    assert len(ctx.report.sections["Would copy"]) == 2 and not (dst / "notes.txt").exists()

    cfg.data["drive_sync"]["destination"] = str(tmp_path / "G" / "My Drive")
    with pytest.raises(RuntimeError, match="Google Drive for desktop"):
        run_drive_sync(make_ctx("drive-sync"))


def test_scheduled_check_is_silent_when_nothing_changed(folders, cfg, tmp_path):
    import yaml

    from origincrest_agent.cli import main

    config = tmp_path / "config.yaml"
    config.write_text(yaml.safe_dump(cfg.data), encoding="utf-8")
    reports = cfg.data_dir / "reports"
    assert main(["--config", str(config), "drive-sync"]) == 0
    first = sorted(p.name for p in reports.glob("20*_drive-sync.md"))
    assert len(first) == 1  # files were copied: one report
    assert main(["--config", str(config), "drive-sync"]) == 0
    assert sorted(p.name for p in reports.glob("20*_drive-sync.md")) == first  # nothing changed: no new report

    cfg.data["drive_sync"]["destination"] = str(tmp_path / "G" / "My Drive")  # Google Drive app not running
    config.write_text(yaml.safe_dump(cfg.data), encoding="utf-8")
    assert main(["--config", str(config), "drive-sync"]) == 1
    assert sorted(p.name for p in reports.glob("20*_drive-sync.md")) == first  # error only in latest_drive-sync.md
    assert "Google Drive folder not found" in (reports / "latest_drive-sync.md").read_text()
