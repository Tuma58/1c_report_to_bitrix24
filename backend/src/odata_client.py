"""Низкоуровневый клиент OData для 1С.

Особенности:
- Basic-auth с логином в кодировке UTF-8 (в cp1251 сервер отдаёт 401).
  Заголовок Authorization формируется вручную через base64.
- Кириллица в путях / именах сущностей — percent-encode (urllib.parse.quote).
- Всегда добавляется $format=json.
- Ретраи с таймаутом; при сетевой недоступности — понятное исключение.
"""
from __future__ import annotations

import base64
import time
from typing import Any, Optional
from urllib.parse import quote, urlencode

import requests

try:
    from .config import Settings, settings as default_settings
except ImportError:  # запуск как скрипта
    from config import Settings, settings as default_settings


class ODataError(RuntimeError):
    """Ошибка обращения к OData 1С (сеть, HTTP, парсинг)."""


class ODataUnavailableError(ODataError):
    """OData недоступен (сетевая ошибка / таймаут / порт закрыт)."""


class ODataClient:
    def __init__(self, config: Optional[Settings] = None) -> None:
        self.settings = config or default_settings
        if not self.settings.odata_base_url:
            raise ODataError("ODATA_BASE_URL не задан в .env")
        self._session = requests.Session()
        self._session.headers.update(
            {
                "Authorization": self._basic_auth_header(),
                "Accept": "application/json",
            }
        )

    def _basic_auth_header(self) -> str:
        # КРИТИЧНО: логин:пароль кодируем в UTF-8, не полагаемся на HTTPBasicAuth.
        raw = f"{self.settings.odata_login}:{self.settings.odata_password}"
        token = base64.b64encode(raw.encode("utf-8")).decode("ascii")
        return f"Basic {token}"

    def _build_url(self, entity: str, params: Optional[dict[str, Any]]) -> str:
        # Имя сущности может содержать кириллицу и скобки — percent-encode.
        # safe: не кодируем то, что валидно в сегменте пути OData.
        path = quote(entity, safe="()/'")
        query: dict[str, Any] = {}
        if params:
            query.update(params)
        query["$format"] = "json"
        # urlencode с quote_via=quote корректно кодирует кириллицу в $filter и т.п.
        qs = urlencode(query, quote_via=quote, safe="() ',$")
        return f"{self.settings.odata_base_url}{path}?{qs}"

    def get(self, entity: str, params: Optional[dict[str, Any]] = None) -> list[dict]:
        """GET сущности. Возвращает список value."""
        url = self._build_url(entity, params)
        data = self._request(url)
        value = data.get("value")
        if value is None:
            # одиночная сущность без value
            return [data] if data else []
        if not isinstance(value, list):
            raise ODataError(f"Неожиданный формат value от {entity}")
        return value

    def metadata_entityset_count(self) -> int:
        """GET $metadata, возвращает число вхождений <EntitySet."""
        url = f"{self.settings.odata_base_url}{quote('$metadata', safe='$')}"
        text = self._request_text(url)
        return text.count("<EntitySet")

    # --- внутреннее ---

    def _request(self, url: str) -> dict:
        resp = self._perform(url)
        try:
            return resp.json()
        except ValueError as exc:
            raise ODataError(
                f"Не удалось разобрать JSON-ответ ({url}): {exc}"
            ) from exc

    def _request_text(self, url: str) -> str:
        return self._perform(url).text

    def _perform(self, url: str) -> requests.Response:
        last_exc: Optional[Exception] = None
        attempts = max(1, self.settings.retries)
        for attempt in range(1, attempts + 1):
            try:
                resp = self._session.get(url, timeout=self.settings.timeout)
            except (requests.ConnectionError, requests.Timeout) as exc:
                last_exc = exc
                if attempt < attempts:
                    time.sleep(min(2 ** (attempt - 1), 5))
                    continue
                raise ODataUnavailableError(
                    "OData недоступен (сеть/таймаут/порт закрыт) после "
                    f"{attempts} попыток: {exc}"
                ) from exc

            if resp.status_code == 401:
                raise ODataError(
                    "401 Unauthorized. Проверьте логин/пароль и кодировку логина "
                    "(должна быть UTF-8)."
                )
            if resp.status_code >= 500:
                last_exc = ODataError(f"HTTP {resp.status_code} от сервера")
                if attempt < attempts:
                    time.sleep(min(2 ** (attempt - 1), 5))
                    continue
            if resp.status_code >= 400:
                raise ODataError(
                    f"HTTP {resp.status_code}: {resp.text[:300]}"
                )
            return resp

        if last_exc:
            raise ODataError(str(last_exc))
        raise ODataError("Неизвестная ошибка запроса к OData")


# --- хелперы для OData-параметров ---

def build_params(
    filter: Optional[str] = None,
    top: Optional[int] = None,
    select: Optional[str] = None,
    orderby: Optional[str] = None,
) -> dict[str, Any]:
    """Собирает словарь OData-параметров ($filter, $top, $select, $orderby)."""
    params: dict[str, Any] = {}
    if filter:
        params["$filter"] = filter
    if top is not None:
        params["$top"] = top
    if select:
        params["$select"] = select
    if orderby:
        params["$orderby"] = orderby
    return params
