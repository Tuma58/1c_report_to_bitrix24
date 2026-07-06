"""MCP-сервер для подключения OpenAI GPT API как агента.

Предоставляет инструменты:
- gpt_chat  — отправить запрос в GPT и получить ответ
- gpt_code  — генерация/анализ кода через GPT
- gpt_analyze — анализ текста/документа через GPT

Запуск:
    python server.py

Конфигурация — переменная окружения OPENAI_API_KEY.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI

# Загружаем .env из backend/
_ENV_PATH = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(_ENV_PATH)

# --- Конфигурация ---
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-4o")
DEFAULT_MAX_TOKENS = int(os.getenv("OPENAI_MAX_TOKENS", "4096"))
DEFAULT_TEMPERATURE = float(os.getenv("OPENAI_TEMPERATURE", "0.7"))

if not OPENAI_API_KEY:
    print("⚠️  OPENAI_API_KEY не задан в .env", file=sys.stderr)
    # Не падаем — сервер запустится, но инструменты вернут ошибку


def _get_client() -> OpenAI:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY не задан. Добавьте в backend/.env: OPENAI_API_KEY=sk-...")
    return OpenAI(api_key=OPENAI_API_KEY)


# ============================================================
# MCP Server (ручная реализация stdio JSON-RPC)
# ============================================================

def _log(msg: str) -> None:
    """Лог в stderr (не ломает stdio-протокол MCP)."""
    print(f"[gpt-server] {msg}", file=sys.stderr, flush=True)


def _send(data: dict[str, Any]) -> None:
    """Отправить JSON-RPC сообщение в stdout."""
    sys.stdout.write(json.dumps(data, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _read() -> dict[str, Any] | None:
    """Прочитать одно JSON-RPC сообщение из stdin."""
    line = sys.stdin.readline()
    if not line:
        return None
    return json.loads(line)


# --- Инструменты ---

TOOLS = [
    {
        "name": "gpt_chat",
        "description": "Отправить запрос в GPT и получить ответ. Используй для: общих вопросов, объяснений, brainstorm, перевода, суммаризации.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "prompt": {
                    "type": "string",
                    "description": "Текст запроса к GPT",
                },
                "system": {
                    "type": "string",
                    "description": "Системный промпт (роль/контекст для GPT)",
                },
                "temperature": {
                    "type": "number",
                    "description": f"Температура (0–2), по умолчанию {DEFAULT_TEMPERATURE}",
                },
            },
            "required": ["prompt"],
        },
    },
    {
        "name": "gpt_code",
        "description": "Генерация, рефакторинг или анализ кода через GPT. Используй для: написать функцию, найти баг, объяснить код, предложить улучшения.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "task": {
                    "type": "string",
                    "description": "Описание задачи: 'напиши функцию', 'найди ошибку', 'отрефактори'",
                },
                "code": {
                    "type": "string",
                    "description": "Исходный код (если нужно проанализировать/исправить)",
                },
                "language": {
                    "type": "string",
                    "description": "Язык программирования: python, javascript, go, etc.",
                },
            },
            "required": ["task"],
        },
    },
    {
        "name": "gpt_analyze",
        "description": "Анализ текста или документа через GPT. Используй для: анализ ТЗ, требований, логов, выделение сути из большого текста.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "content": {
                    "type": "string",
                    "description": "Текст для анализа",
                },
                "question": {
                    "type": "string",
                    "description": "Что нужно выяснить/извлечь из текста",
                },
            },
            "required": ["content", "question"],
        },
    },
]


def _call_gpt(messages: list[dict], temperature: float = DEFAULT_TEMPERATURE) -> str:
    """Вызов OpenAI Chat Completion."""
    client = _get_client()
    response = client.chat.completions.create(
        model=DEFAULT_MODEL,
        messages=messages,
        max_tokens=DEFAULT_MAX_TOKENS,
        temperature=temperature,
    )
    return response.choices[0].message.content or ""


def handle_tool_call(name: str, arguments: dict[str, Any]) -> list[dict]:
    """Обработчик вызова инструмента."""
    _log(f"tool call: {name}({json.dumps(arguments, ensure_ascii=False)[:200]})")

    if name == "gpt_chat":
        prompt = arguments.get("prompt", "")
        system = arguments.get("system") or "Ты — полезный ассистент. Отвечай кратко и по делу."
        temperature = arguments.get("temperature", DEFAULT_TEMPERATURE)
        messages = [
            {"role": "system", "content": system},
            {"role": "user", "content": prompt},
        ]
        result = _call_gpt(messages, temperature)
        return [{"type": "text", "text": result}]

    elif name == "gpt_code":
        task = arguments.get("task", "")
        code = arguments.get("code", "")
        language = arguments.get("language", "")
        lang_hint = f" (язык: {language})" if language else ""
        system = f"Ты — эксперт-программист{lang_hint}. Пиши чистый, рабочий код. Объясняй кратко."
        user_msg = f"Задача: {task}"
        if code:
            user_msg += f"\n\nКод:\n```\n{code}\n```"
        messages = [
            {"role": "system", "content": system},
            {"role": "user", "content": user_msg},
        ]
        result = _call_gpt(messages, temperature=0.3)
        return [{"type": "text", "text": result}]

    elif name == "gpt_analyze":
        content = arguments.get("content", "")
        question = arguments.get("question", "")
        system = "Ты — аналитик. Извлекай суть, находи ключевые моменты. Отвечай структурированно."
        messages = [
            {"role": "system", "content": system},
            {"role": "user", "content": f"Вопрос: {question}\n\nТекст для анализа:\n{content}"},
        ]
        result = _call_gpt(messages, temperature=0.5)
        return [{"type": "text", "text": result}]

    else:
        return [{"type": "text", "text": f"Неизвестный инструмент: {name}"}]


# ============================================================
# Основной цикл MCP (stdio JSON-RPC)
# ============================================================

def run() -> None:
    _log(f"Starting GPT MCP server (model={DEFAULT_MODEL}, temp={DEFAULT_TEMPERATURE})")

    for line in sys.stdin:
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            continue

        msg_id = msg.get("id")
        method = msg.get("method", "")

        if method == "initialize":
            _send({
                "jsonrpc": "2.0",
                "id": msg_id,
                "result": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {"tools": {}},
                    "serverInfo": {
                        "name": "gpt-mcp-server",
                        "version": "1.0.0",
                    },
                },
            })

        elif method == "notifications/initialized":
            pass  # не отвечаем на notifications

        elif method == "tools/list":
            _send({
                "jsonrpc": "2.0",
                "id": msg_id,
                "result": {"tools": TOOLS},
            })

        elif method == "tools/call":
            params = msg.get("params", {})
            tool_name = params.get("name", "")
            tool_args = params.get("arguments", {})
            try:
                content = handle_tool_call(tool_name, tool_args)
                _send({
                    "jsonrpc": "2.0",
                    "id": msg_id,
                    "result": {"content": content},
                })
            except Exception as e:
                _log(f"ERROR: {e}")
                _send({
                    "jsonrpc": "2.0",
                    "id": msg_id,
                    "error": {"code": -32000, "message": str(e)},
                })

        elif method == "ping":
            _send({"jsonrpc": "2.0", "id": msg_id, "result": {}})

        else:
            _send({
                "jsonrpc": "2.0",
                "id": msg_id,
                "error": {"code": -32601, "message": f"Method not found: {method}"},
            })


if __name__ == "__main__":
    run()
