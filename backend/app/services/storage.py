from pathlib import Path
from uuid import uuid4

from app.core.config import Settings, get_settings


class LocalStorage:
    def __init__(self, settings: Settings | None = None) -> None:
        self.settings = settings or get_settings()
        self.root = Path(self.settings.local_storage_path)
        self.root.mkdir(parents=True, exist_ok=True)
        (self.root / "uploads").mkdir(exist_ok=True)
        (self.root / "generated").mkdir(exist_ok=True)
        (self.root / "specs").mkdir(exist_ok=True)

    def save_upload(self, filename: str, content: bytes) -> str:
        suffix = Path(filename).suffix.lower() or ".csv"
        path = self.root / "uploads" / f"{uuid4()}{suffix}"
        path.write_bytes(content)
        return str(path.relative_to(self.root))

    def save_generated_file(self, file_id: str, content: bytes) -> str:
        path = self.root / "generated" / f"{file_id}.xlsx"
        path.write_bytes(content)
        return str(path.relative_to(self.root))

    def read(self, storage_path: str) -> bytes:
        return self._resolve(storage_path).read_bytes()

    def path_for(self, storage_path: str) -> Path:
        return self._resolve(storage_path)

    def delete(self, storage_path: str) -> None:
        path = self._resolve(storage_path)
        if path.exists() and path.is_file():
            path.unlink()

    def _resolve(self, storage_path: str) -> Path:
        root = self.root.resolve()
        path = (self.root / storage_path).resolve()
        if root != path and root not in path.parents:
            raise ValueError("Storage path escapes configured root")
        return path

