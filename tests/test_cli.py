import logging

from origincrest_agent import cli
from origincrest_agent.utils import SharedLogFile


def test_log_keeps_writing_when_another_run_blocks_rotation(tmp_path, monkeypatch):
    path = tmp_path / "agent.log"
    handler = SharedLogFile(path, maxBytes=50, backupCount=2, encoding="utf-8")

    def locked(source, dest):  # what Windows does while another run has the log open
        raise PermissionError(13, "The process cannot access the file")

    monkeypatch.setattr(handler, "rotate", locked)
    logger = logging.getLogger("origincrest.test_shared_log")
    logger.addHandler(handler)
    try:
        for i in range(5):
            logger.warning("line %d %s", i, "x" * 30)
    finally:
        logger.removeHandler(handler)
        handler.close()
    text = path.read_text(encoding="utf-8")
    assert all(f"line {i}" in text for i in range(5))


def test_unexpected_error_is_logged_not_lost(tmp_path, monkeypatch, caplog):
    config = tmp_path / "config.yaml"
    config.write_text(f"data_dir: '{tmp_path / 'data'}'\n", encoding="utf-8")
    monkeypatch.setattr(cli, "setup_logging", lambda *a, **k: None)

    def crash(cfg, args):
        raise RuntimeError("disk on fire")

    monkeypatch.setattr(cli, "run_workflow", crash)
    with caplog.at_level(logging.ERROR, logger="origincrest"):
        assert cli.main(["--config", str(config), "followup"]) == 1
    assert "'followup' stopped with an unexpected error" in caplog.text
    assert "disk on fire" in caplog.text


def test_config_layout_mistake_is_explained_and_logged(tmp_path, monkeypatch):
    monkeypatch.setenv("HOME", str(tmp_path))
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    config = tmp_path / "config.yaml"
    config.write_text("discovery:\n  profiles:\n    rfq:\n      categories: []\n     other_tab: x\n", encoding="utf-8")
    assert cli.main(["--config", str(config), "followup"]) == 2
    logged = next(tmp_path.rglob("agent.log")).read_text(encoding="utf-8")
    assert "ERROR" in logged and "layout mistake" in logged
