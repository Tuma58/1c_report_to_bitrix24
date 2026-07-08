"""Отправка созданных Excel-отчётов по SMTP."""
from __future__ import annotations

import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, Optional

try:
    from .config import Settings, settings as default_settings
except ImportError:  # запуск как скрипта
    from config import Settings, settings as default_settings


class EmailError(RuntimeError):
    """Ошибка SMTP-отправки отчётов."""


class EmailSender:
    def __init__(self, config: Optional[Settings] = None) -> None:
        self.settings = config or default_settings
        if not self.settings.email_smtp_host:
            raise EmailError("EMAIL_SMTP_HOST не задан в .env")
        if not self.settings.email_from:
            raise EmailError("EMAIL_FROM не задан в .env")
        if not self.settings.email_to:
            raise EmailError("EMAIL_TO не задан в .env")

    @staticmethod
    def _recipients(raw: str) -> list[str]:
        return [item.strip() for item in raw.replace(";", ",").split(",") if item.strip()]

    def send_files(self, files: Iterable[Path], subject: str, body: str) -> None:
        recipients = self._recipients(self.settings.email_to)
        if not recipients:
            raise EmailError("EMAIL_TO не содержит адресов получателей")

        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = self.settings.email_from
        msg["To"] = ", ".join(recipients)
        msg.set_content(body)

        for file_path in files:
            path = Path(file_path)
            msg.add_attachment(
                path.read_bytes(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=path.name,
            )

        try:
            if self.settings.email_use_ssl:
                with smtplib.SMTP_SSL(
                    self.settings.email_smtp_host,
                    self.settings.email_smtp_port,
                    timeout=60,
                ) as smtp:
                    self._send(smtp, msg, recipients)
            else:
                with smtplib.SMTP(
                    self.settings.email_smtp_host,
                    self.settings.email_smtp_port,
                    timeout=60,
                ) as smtp:
                    if self.settings.email_use_tls:
                        smtp.starttls()
                    self._send(smtp, msg, recipients)
        except OSError as exc:
            raise EmailError(f"SMTP недоступен: {exc}") from exc
        except smtplib.SMTPException as exc:
            raise EmailError(f"SMTP ошибка: {exc}") from exc

    def _send(self, smtp: smtplib.SMTP, msg: EmailMessage, recipients: list[str]) -> None:
        if self.settings.email_smtp_login:
            smtp.login(self.settings.email_smtp_login, self.settings.email_smtp_password)
        smtp.send_message(msg, from_addr=self.settings.email_from, to_addrs=recipients)
