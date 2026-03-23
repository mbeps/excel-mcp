from logging import Logger

from mcp.server.fastmcp import FastMCP

from mcp_server.models import AnalysisResult, EchoRequest, SystemMetrics
from mcp_server.tools.examples import render_echo
from mcp_server.tools.system import calculate_health
from mcp_server.utils.logger import configure_logging

mcp: FastMCP = FastMCP("mcp-template", dependencies=["pydantic"])
logger: Logger = configure_logging("mcp_server.main")


@mcp.tool()
def analyze_metrics(metrics: SystemMetrics) -> AnalysisResult:
    """Analyze provided system metrics and return a health assessment."""
    logger.info("Received analyze_metrics request", extra={"cpu_percent": metrics.cpu_percent})
    return calculate_health(metrics)


@mcp.tool()
def echo(payload: EchoRequest) -> str:
    """Echo the provided message with optional formatting tweaks."""
    return render_echo(payload)


@mcp.resource("template://config/defaults")
def get_default_config() -> str:
    """Return a starter configuration payload for new MCP servers."""
    logger.info("Serving default config resource")
    return (
        "{\n"
        '  "project_name": "mcp-template",\n'
        '  "description": "Replace with your server description."\n'
        "}\n"
    )


def run() -> None:
    """Entrypoint for launching the MCP server."""
    try:
        mcp.run()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception:
        logger.critical("Fatal server error", exc_info=True)
        raise


if __name__ == "__main__":
    run()
