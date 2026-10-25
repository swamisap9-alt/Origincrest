"""Claude API access.

Two call shapes:
* `research()` - Claude searches/fetches the web (server-side tools) and hands
  its findings back through a strict "record" tool, so the result always
  matches our schema.
* `structured()` - no tools; the reply is constrained to a JSON schema
  (used for writing email drafts).
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any

import anthropic

from .config import ConfigError
from .utils import log

FALLBACK_BETA = "server-side-fallback-2026-07-01"


class AgentError(Exception):
    pass


class AgentRefusal(AgentError):
    pass


@dataclass
class ResearchResult:
    data: dict[str, Any]
    # (url, text) of every page Claude fetched - used to verify contact details.
    pages: list[tuple[str, str]] = field(default_factory=list)
    urls: set[str] = field(default_factory=set)


class ClaudeAgent:
    def __init__(self, cfg):
        c = cfg["claude"]
        api_key = cfg.secret(c["api_key_env"])
        if not api_key:
            raise ConfigError(
                f"{c['api_key_env']} is not set. Add it to the .env file next to config.yaml "
                "(see README: 'Anthropic API key')."
            )
        self.client = anthropic.Anthropic(api_key=api_key, timeout=c["timeout_seconds"], max_retries=4)
        self.model = c["model"]
        self.max_tokens = c["max_tokens"]
        self.use_fallbacks = c["use_fallbacks"]
        self.search_tool = c["web_search_tool"]
        self.fetch_tool = c["web_fetch_tool"]
        self.user_location = c["user_location"]
        self.usage = {"requests": 0, "input_tokens": 0, "output_tokens": 0, "cache_read_input_tokens": 0,
                      "web_search_requests": 0, "web_fetch_requests": 0}

    def _call(self, *, system: str, messages: list, effort: str, tools: list | None = None,
              output_format: dict | None = None):
        output_config: dict[str, Any] = {"effort": effort}
        if output_format:
            output_config["format"] = output_format
        kwargs: dict[str, Any] = dict(
            model=self.model,
            max_tokens=self.max_tokens,
            system=system,
            messages=messages,
            thinking={"type": "adaptive"},
            output_config=output_config,
        )
        if tools:
            kwargs["tools"] = tools
        if self.use_fallbacks:
            kwargs["betas"] = [FALLBACK_BETA]
            kwargs["fallbacks"] = "default"
        with self.client.beta.messages.stream(**kwargs) as stream:
            response = stream.get_final_message()
        self._add_usage(response)
        if response.stop_reason == "refusal":
            details = getattr(response, "stop_details", None)
            raise AgentRefusal(f"Claude declined the request ({getattr(details, 'category', None) or 'no category'})")
        if response.stop_reason == "max_tokens":
            raise AgentError("Claude hit max_tokens; raise claude.max_tokens in config.yaml")
        return response

    def _add_usage(self, response) -> None:
        u = response.usage
        self.usage["requests"] += 1
        self.usage["input_tokens"] += (u.input_tokens or 0) + (u.cache_creation_input_tokens or 0)
        self.usage["output_tokens"] += u.output_tokens or 0
        self.usage["cache_read_input_tokens"] += u.cache_read_input_tokens or 0
        if u.server_tool_use:
            self.usage["web_search_requests"] += u.server_tool_use.web_search_requests or 0
            self.usage["web_fetch_requests"] += u.server_tool_use.web_fetch_requests or 0

    def research(self, *, system: str, prompt: str, record_tool: dict, effort: str,
                 web_searches: int, web_fetches: int) -> ResearchResult:
        tools: list[dict] = []
        if web_searches:
            tools.append({"type": self.search_tool, "name": "web_search", "max_uses": web_searches,
                          "user_location": self.user_location})
        if web_fetches:
            tools.append({"type": self.fetch_tool, "name": "web_fetch", "max_uses": web_fetches,
                          "max_content_tokens": 12000})
        tools.append(record_tool)
        messages: list[dict] = [{"role": "user", "content": prompt}]
        result = ResearchResult(data={})
        nudges = 0
        for _ in range(12):
            response = self._call(system=system, messages=messages, effort=effort, tools=tools)
            _collect_evidence(response, result)
            if response.stop_reason == "tool_use":
                for block in response.content:
                    if block.type == "tool_use" and block.name == record_tool["name"]:
                        result.data = dict(block.input)
                        return result
            messages.append({"role": "assistant", "content": response.content})
            if response.stop_reason == "pause_turn":
                continue  # server-side tool loop paused; re-send to let it resume
            nudges += 1
            if nudges > 2:
                break
            messages.append({"role": "user", "content": (
                f"Please call the {record_tool['name']} tool now with what you found "
                "(an empty list is fine if nothing qualified).")})
        raise AgentError(f"Claude did not call {record_tool['name']}")

    def structured(self, *, system: str, prompt: str, schema: dict, effort: str) -> dict[str, Any]:
        response = self._call(
            system=system,
            messages=[{"role": "user", "content": prompt}],
            effort=effort,
            output_format={"type": "json_schema", "schema": schema},
        )
        text = next((b.text for b in response.content if b.type == "text"), "")
        try:
            return json.loads(text)
        except json.JSONDecodeError as exc:
            raise AgentError(f"Claude returned invalid JSON: {text[:200]}") from exc


def _collect_evidence(response, result: ResearchResult) -> None:
    for block in response.content:
        btype = getattr(block, "type", "")
        if btype == "web_fetch_tool_result":
            content = getattr(block, "content", None)
            url = getattr(content, "url", None)
            source = getattr(getattr(content, "content", None), "source", None)
            if url:
                result.urls.add(url)
                if getattr(source, "type", "") == "text" and getattr(source, "data", None):
                    result.pages.append((url, source.data))
        elif btype == "web_search_tool_result":
            content = getattr(block, "content", None)
            if isinstance(content, list):
                result.urls.update(r.url for r in content if getattr(r, "url", None))
        elif btype == "text":
            for citation in getattr(block, "citations", None) or []:
                if getattr(citation, "url", None):
                    result.urls.add(citation.url)


def strict_tool(name: str, description: str, schema: dict) -> dict:
    return {"name": name, "description": description, "strict": True, "input_schema": schema}


def obj(properties: dict[str, Any]) -> dict[str, Any]:
    """JSON-schema object where every property is required (unknowns are empty strings)."""
    return {"type": "object", "properties": properties, "required": list(properties), "additionalProperties": False}


def log_usage(agent: "ClaudeAgent | None") -> dict[str, int]:
    if agent is None:
        return {}
    log.info(
        "Claude usage: %d requests, %d input tokens (+%d cached), %d output tokens, %d web searches, %d fetches",
        agent.usage["requests"], agent.usage["input_tokens"], agent.usage["cache_read_input_tokens"],
        agent.usage["output_tokens"], agent.usage["web_search_requests"], agent.usage["web_fetch_requests"],
    )
    return dict(agent.usage)
