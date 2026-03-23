from logging import Logger

from mcp_server.models import AnalysisResult, SystemMetrics
from mcp_server.utils.logger import configure_logging

logger: Logger = configure_logging(__name__)


def calculate_health(metrics: SystemMetrics) -> AnalysisResult:
    """Determine system health based on CPU and memory load.

    Args:
        metrics: Snapshot of resource usage for the host under evaluation.

    Returns:
        Health classification with optional remediation guidance.
    """
    logger.info(
        "Analyzing system metrics", extra={"cpu_percent": metrics.cpu_percent}
    )

    if metrics.cpu_percent > 90.0 or metrics.memory_gb > 32.0:
        return AnalysisResult(
            status="critical",
            recommendation="Immediate scale-up or process termination required.",
        )
    if metrics.cpu_percent > 70.0:
        return AnalysisResult(
            status="warning", recommendation="Monitor system load closely."
        )

    return AnalysisResult(status="healthy", recommendation=None)
