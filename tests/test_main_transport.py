"""Transport-mode tests for mcp_server.main.run."""

from __future__ import annotations

from argparse import Namespace
from unittest.mock import Mock

import pytest

import mcp_server.main as main_module


def _mock_parsed_args(monkeypatch: pytest.MonkeyPatch, args: Namespace) -> None:
    """Patch argparse.ArgumentParser.parse_args to return a fixed namespace."""

    parse_args_mock = Mock(return_value=args)
    monkeypatch.setattr("argparse.ArgumentParser.parse_args", parse_args_mock)


def test_run_defaults_to_stdio_calls_noarg_mcp_run(monkeypatch: pytest.MonkeyPatch) -> None:
    """Default stdio mode must call mcp.run() with no kwargs."""

    run_mock = Mock()
    monkeypatch.setattr(main_module, "mcp", Mock(run=run_mock))
    _mock_parsed_args(monkeypatch, Namespace(transport="stdio", host="127.0.0.1", port=8000))

    main_module.run()

    run_mock.assert_called_once_with()


def test_run_http_mode_calls_streamable_http(monkeypatch: pytest.MonkeyPatch) -> None:
    """HTTP mode should pass streamable-http transport and host/port."""

    run_mock = Mock()
    settings_mock = Mock(host="127.0.0.1", port=8000)
    monkeypatch.setattr(main_module, "mcp", Mock(run=run_mock, settings=settings_mock))
    _mock_parsed_args(monkeypatch, Namespace(transport="http", host="0.0.0.0", port=9000))

    main_module.run()

    assert settings_mock.host == "0.0.0.0"
    assert settings_mock.port == 9000
    run_mock.assert_called_once_with(transport="streamable-http")


def test_run_handles_keyboard_interrupt(monkeypatch: pytest.MonkeyPatch) -> None:
    """KeyboardInterrupt should be swallowed and logged."""

    run_mock = Mock(side_effect=KeyboardInterrupt())
    logger_mock = Mock()
    monkeypatch.setattr(main_module, "mcp", Mock(run=run_mock))
    monkeypatch.setattr(main_module, "logger", logger_mock)
    _mock_parsed_args(monkeypatch, Namespace(transport="stdio", host="127.0.0.1", port=8000))

    main_module.run()

    logger_mock.info.assert_called()


def test_run_reraises_unexpected_exception(monkeypatch: pytest.MonkeyPatch) -> None:
    """Unexpected exceptions should be logged and re-raised."""

    run_mock = Mock(side_effect=RuntimeError("boom"))
    logger_mock = Mock()
    monkeypatch.setattr(main_module, "mcp", Mock(run=run_mock))
    monkeypatch.setattr(main_module, "logger", logger_mock)
    _mock_parsed_args(monkeypatch, Namespace(transport="stdio", host="127.0.0.1", port=8000))

    with pytest.raises(RuntimeError, match="boom"):
        main_module.run()

    logger_mock.critical.assert_called_once()
