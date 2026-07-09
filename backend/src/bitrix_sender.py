"""Отправка созданных отчётов в чат Bitrix24 через REST webhook.

Файлы отправляются как вложения самого чата: сначала берётся IM-папка диалога,
затем xlsx загружается в неё и прикрепляется к сообщению через im.disk.file.commit.
Публичный URL на VPS и отдельная папка общего Диска для этого не нужны.
"""
from __future__ import annotations

import base64
import secrets
from datetime import datetime
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

    def _chat_folder_id(self) -> int:
        data = self._post("im.disk.folder.get", {"DIALOG_ID": self.chat_id})
        result = data.get("result")
        if not isinstance(result, dict) or not result.get("ID"):
            raise BitrixError("Bitrix24 не вернул ID папки чата")
        return int(result["ID"])

    @staticmethod
    def _unique_chat_filename(path: Path) -> str:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        suffix = secrets.token_hex(2)
        return f"{path.stem}_{stamp}_{suffix}{path.suffix}"

    def upload_file(self, path: Path) -> dict:
        """Загружает файл в IM-папку чата и возвращает result."""
        file_path = Path(path)
        folder_id = self._chat_folder_id()
        content = base64.b64encode(file_path.read_bytes()).decode("ascii")
        data = self._post(
            "disk.folder.uploadfile",
            {
                "id": folder_id,
                "data": {"NAME": self._unique_chat_filename(file_path)},
                "fileContent": content,
            },
        )
        return data.get("result", {})

    def _commit_file(self, upload_result: dict, message: str) -> dict:
        file_id = upload_result.get("ID")
        if not file_id:
            raise BitrixError("Bitrix24 не вернул ID загруженного файла")
        return self._post(
            "im.disk.file.commit",
            {
                "DIALOG_ID": self.chat_id,
                "FILE_ID": int(file_id),
                "MESSAGE": message,
            },
        )

    def send_files(self, files: Iterable[Path], title: str) -> None:
        """Отправляет файлы в чат как вложения."""
        for index, file_path in enumerate(files):
            result = self.upload_file(Path(file_path))
            message = title if index == 0 else ""
            self._commit_file(result, message)
