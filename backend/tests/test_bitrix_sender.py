"""Smoke-test for Bitrix24 chat attachment flow.

No network calls: the REST layer is replaced with an in-memory fake.
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from bitrix_sender import BitrixSender  # noqa: E402
from config import Settings  # noqa: E402


class FakeBitrixSender(BitrixSender):
    def __init__(self) -> None:
        settings = Settings(
            odata_base_url="",
            odata_login="",
            odata_password="",
            timeout=30,
            retries=3,
            bitrix_webhook_url="https://example.bitrix24.ru/rest/1/token/",
            bitrix_chat_id="chat226489",
            email_smtp_host="",
            email_smtp_port=587,
            email_smtp_login="",
            email_smtp_password="",
            email_from="",
            email_to="",
            email_use_tls=True,
            email_use_ssl=False,
        )
        super().__init__(settings)
        self.calls: list[tuple[str, dict]] = []

    def _post(self, method: str, payload: dict) -> dict:
        self.calls.append((method, payload))
        if method == "im.disk.folder.get":
            return {"result": {"ID": 1815267}}
        if method == "disk.folder.uploadfile":
            return {"result": {"ID": 1815269, "NAME": payload["data"]["NAME"]}}
        if method == "im.disk.file.commit":
            return {"result": {"MESSAGE_ID": 9012355}}
        raise AssertionError(f"unexpected method: {method}")


def main() -> int:
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "report.xlsx"
        path.write_bytes(b"test xlsx bytes")

        sender = FakeBitrixSender()
        sender.send_files([path], "Test report")

    methods = [method for method, _payload in sender.calls]
    assert methods == ["im.disk.folder.get", "disk.folder.uploadfile", "im.disk.file.commit"], methods

    upload_payload = sender.calls[1][1]
    assert upload_payload["id"] == 1815267
    assert upload_payload["data"]["NAME"].startswith("report_")
    assert upload_payload["data"]["NAME"].endswith(".xlsx")
    assert upload_payload["fileContent"]

    commit_payload = sender.calls[2][1]
    assert commit_payload == {
        "DIALOG_ID": "chat226489",
        "FILE_ID": 1815269,
        "MESSAGE": "Test report",
    }

    print("bitrix sender chat attachment smoke ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
