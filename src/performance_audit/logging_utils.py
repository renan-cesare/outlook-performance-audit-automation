from datetime import datetime
from pathlib import Path


class Logger:
    def __init__(self, log_path: str | None = None):
        self.log_path = Path(log_path) if log_path else None
        if self.log_path:
            self.log_path.parent.mkdir(parents=True, exist_ok=True)

    def info(self, msg: str):
        self._write("[INFO] " + msg)

    def warn(self, msg: str):
        self._write("[WARN] " + msg)

    def error(self, msg: str):
        self._write("[ERRO] " + msg)

    def _write(self, msg: str):
        line = f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}"
        print(line)
        if self.log_path:
            try:
                existing = self.log_path.read_text(encoding="utf-8") if self.log_path.exists() else ""
                self.log_path.write_text(existing + line + "\n", encoding="utf-8")
            except Exception:
                pass

