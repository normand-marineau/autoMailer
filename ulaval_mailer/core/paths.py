from __future__ import annotations

from pathlib import Path

def ensure_logs_dir(base: Path | None = None) -> Path:
    root = base or Path.cwd()
    logs = root / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    return logs
