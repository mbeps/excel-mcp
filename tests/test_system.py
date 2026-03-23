from mcp_server.models import AnalysisResult, SystemMetrics
from mcp_server.tools.system import calculate_health


def test_calculate_health_critical_cpu() -> None:
    metrics = SystemMetrics(cpu_percent=95.0, memory_gb=8.0, process_count=200)
    result: AnalysisResult = calculate_health(metrics)
    assert result.status == "critical"
    assert "Immediate" in (result.recommendation or "")


def test_calculate_health_warning_cpu() -> None:
    metrics = SystemMetrics(cpu_percent=75.0, memory_gb=8.0, process_count=200)
    result: AnalysisResult = calculate_health(metrics)
    assert result.status == "warning"
    assert "Monitor" in (result.recommendation or "")


def test_calculate_health_healthy() -> None:
    metrics = SystemMetrics(cpu_percent=30.0, memory_gb=4.0, process_count=120)
    result: AnalysisResult = calculate_health(metrics)
    assert result.status == "healthy"
    assert result.recommendation is None
