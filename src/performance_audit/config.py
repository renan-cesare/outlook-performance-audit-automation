import json
from pathlib import Path


class Config:
    def __init__(self, data: dict):
        self.data = data

    @classmethod
    def load(cls, path: str) -> "Config":
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(f"Config não encontrada: {p.resolve()}")
        return cls(json.loads(p.read_text(encoding="utf-8")))

    def get(self, dotted: str, default=None):
        cur = self.data
        for part in dotted.split("."):
            if isinstance(cur, dict) and part in cur:
                cur = cur[part]
            else:
                return default
        return cur

