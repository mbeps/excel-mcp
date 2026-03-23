from logging import Logger

from mcp_server.models import EchoRequest
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def render_echo(payload: EchoRequest) -> str:
    """Return the provided message with optional formatting and repetition."""
    logger.info("Rendering echo payload", extra={"repeat": payload.repeat})
    text: str = payload.message.upper() if payload.uppercase else payload.message
    return " ".join([text] * payload.repeat)
