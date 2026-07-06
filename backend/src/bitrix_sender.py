"""Отправка созданных отчётов в чат Bitrix24 через REST webhook.

Файлы загружаются в папку Диска Bitrix24 (`BITRIX_DISK_FOLDER_ID`), после чего
в чат отправляется сообщение со ссылками на загруженные xlsx. Такой вариант
не требует публичного URL на VPS.
"""
from __future__ import annotations

import base64
from pathlib import Path
from typing import Iterable, Optional

import requests

try:
    from .config import Settings, settings as default_settings
except ImportError:  # запуск как скрипта
    from config import Settings, settings as default_settings


class BitrixError(RuntimeError):
    """Ошибка REST-вызова Bitrix24."""


class BitrixSender:
    def __init__(self, config: Optional[Settings] = None) -> None:
        self.settings = config or default_settings
        self.webhook_url = self.settings.bitrix_webhook_url.rstrip("/")
        self.chat_id = self.settings.bitrix_chat_id
        self.disk_folder_id = self.settings.bitrix_disk_folder_id
        if not self.webhook_url:
            raise BitrixError("BITRIX_WEBHOOK_URL не задан в .env")
        if not self.chat_id:
            raise BitrixError("BITRIX_CHAT_ID не задан в .env")

    def _method_url(self, method: str) -> str:
        return f"{self.webhook_url}/{method}.json"

    def _post(self, method: str, payload: dict) -> dict:
        try:
            response = requests.post(self._method_url(method), json=payload, timeout=60)
        except requests.RequestException as exc:
            raise BitrixError(f"Bitrix24 недоступен: {exc}") from exc

        try:
            data = response.json()
        except ValueError as exc:
            raise BitrixError(
                f"Bitrix24 вернул не-JSON ответ HTTP {response.status_code}: "
                f"{response.text[:300]}"
            ) from exc

        if response.status_code >= 400 or "error" in data:
            error = data.get("error_description") or data.get("error") or response.text[:300]
            raise BitrixError(f"Bitrix24 {method} failed: {error}")
        return data

    def send_message(self, message: str) -> dict:
        """Отправляет текстовое сообщение в чат."""
        return self._post(
            "im.message.add",
            {
                "DIALOG_ID": self.chat_id,
                "MESSAGE": message,
            },
        )

    def upload_file(self, path: Path) -> dict:
        """Загружает файл в папку Диска Bitrix24 и возвращает result."""
        if not self.disk_folder_id:
            raise BitrixError(
                "BITRIX_DISK_FOLDER_ID не задан: без папки Диска нельзя "
                "загрузить xlsx и отправить ссылку в чат."
            )

        file_path = Path(path)
        content = base64.b64encode(file_path.read_bytes()).decode("ascii")
        data = self._post(
            "disk.folder.uploadfile",
            {
                "id": self.disk_folder_id,
                "data": {"NAME": file_path.name},
                "fileContent": content,
            },
        )
        return data.get("result", {})

    @staticmethod
    def _file_url(upload_result: dict) -> str:
        for key in ("DETAIL_URL", "DOWNLOAD_URL", "VIEW_URL"):
            value = upload_result.get(key)
            if value:
                return str(value)
        return ""

    def send_files(self, files: Iterable[Path], title: str) -> None:
        """Загружает файлы на Диск и отправляет в чат сообщение со ссылками."""
        uploaded: list[tuple[Path, str]] = []
        for file_path in files:
            result = self.upload_file(Path(file_path))
            uploaded.append((Path(file_path), self._file_url(result)))

        lines = [title, ""]
        for file_path, url in uploaded:
            lines.append(f"- {file_path.name}: {url or 'загружено в Диск'}")
        self.send_message("\n".join(lines))
