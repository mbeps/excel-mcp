from typing import Literal, Optional

from pydantic import BaseModel, Field


class SystemMetrics(BaseModel):
    """Snapshot of resource usage used to assess host health."""

    cpu_percent: float = Field(..., description="Current CPU load percentage (0-100).")
    memory_gb: float = Field(..., description="RAM in use in gigabytes.")
    process_count: int = Field(..., description="Number of running processes.")


class AnalysisResult(BaseModel):
    """Structured result for system health analysis."""

    status: Literal["healthy", "warning", "critical"] = Field(
        ..., description="Overall system health rating."
    )
    recommendation: Optional[str] = Field(
        None, description="Actionable advice when health is not healthy."
    )


class EchoRequest(BaseModel):
    """Simple payload for the example echo tool."""

    message: str = Field(..., description="Text to echo back to the caller.")
    repeat: int = Field(
        1,
        ge=1,
        le=5,
        description="Number of times to repeat the message (capped to avoid log spam).",
    )
    uppercase: bool = Field(
        False, description="Return the echoed text in uppercase when true."
    )
