from mcp_server.models import EchoRequest
from mcp_server.tools.examples import render_echo


def test_render_echo_repeats_message() -> None:
    payload = EchoRequest(message="hello", repeat=2, uppercase=False)
    assert render_echo(payload) == "hello hello"


def test_render_echo_applies_uppercase() -> None:
    payload = EchoRequest(message="hi", repeat=1, uppercase=True)
    assert render_echo(payload) == "HI"
