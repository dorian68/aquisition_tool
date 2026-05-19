import csv
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pandas as pd
from pandas.errors import EmptyDataError, ParserError

from app.core.config import Settings, get_settings


ALLOWED_MIME_TYPES = {
    "text/csv",
    "application/csv",
    "application/vnd.ms-excel",
    "text/plain",
    "application/octet-stream",
}


class CSVValidationError(ValueError):
    pass


@dataclass(frozen=True)
class LoadedCSV:
    dataframe: pd.DataFrame
    delimiter: str
    encoding: str
    rows: int
    columns: int


class CSVLoader:
    def __init__(self, settings: Settings | None = None) -> None:
        self.settings = settings or get_settings()

    def load_bytes(self, filename: str, content: bytes, content_type: str | None = None) -> LoadedCSV:
        self._validate_file(filename, content, content_type)
        text, encoding = self._decode(content)
        delimiter = self._detect_delimiter(text)

        try:
            df = pd.read_csv(BytesIO(content), sep=delimiter, encoding=encoding)
        except (EmptyDataError, ParserError, UnicodeDecodeError) as exc:
            raise CSVValidationError("CSV could not be parsed") from exc

        df = df.dropna(how="all")
        if df.empty or len(df.columns) == 0:
            raise CSVValidationError("CSV is empty")

        if len(df) > self.settings.csv_preview_row_limit:
            df = df.head(self.settings.csv_preview_row_limit).copy()

        return LoadedCSV(
            dataframe=df,
            delimiter=delimiter,
            encoding=encoding,
            rows=int(len(df)),
            columns=int(len(df.columns)),
        )

    def _validate_file(self, filename: str, content: bytes, content_type: str | None) -> None:
        if Path(filename).suffix.lower() != ".csv":
            raise CSVValidationError("Only .csv files are accepted")
        if content_type and content_type not in ALLOWED_MIME_TYPES:
            raise CSVValidationError("Invalid CSV MIME type")
        if not content:
            raise CSVValidationError("CSV is empty")
        max_bytes = self.settings.max_upload_mb * 1024 * 1024
        if len(content) > max_bytes:
            raise CSVValidationError(f"CSV exceeds {self.settings.max_upload_mb} MB")

    @staticmethod
    def _decode(content: bytes) -> tuple[str, str]:
        for encoding in ("utf-8-sig", "utf-8", "latin1"):
            try:
                return content.decode(encoding), encoding
            except UnicodeDecodeError:
                continue
        raise CSVValidationError("Unsupported CSV encoding")

    @staticmethod
    def _detect_delimiter(text: str) -> str:
        sample = text[:8192]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            return dialect.delimiter
        except csv.Error:
            counts = {delimiter: sample.count(delimiter) for delimiter in [",", ";", "\t"]}
            return max(counts, key=counts.get) if any(counts.values()) else ","

