"""Drive ClaudeAgent through the real Anthropic SDK with a mocked HTTP transport.

Checks the request we send (model, tools, fallbacks, output_config) is accepted
by the SDK, and that streaming, pause_turn and evidence collection work.
"""

import json

import anthropic
import httpx2
import pytest

from origincrest_agent.llm import ClaudeAgent
from origincrest_agent.prompts import EMAIL_SCHEMA, RECORD_COMPANIES


def sse(events):
    return "".join(f"event: {e['type']}\ndata: {json.dumps(e)}\n\n" for e in events).encode()


def message(blocks, stop_reason):
    events = [{"type": "message_start", "message": {
        "id": "msg_1", "type": "message", "role": "assistant", "model": "claude-opus-5-5", "content": [],
        "stop_reason": None, "stop_sequence": None, "usage": {"input_tokens": 100, "output_tokens": 1}}}]
    for i, (block, partial) in enumerate(blocks):
        events.append({"type": "content_block_start", "index": i, "content_block": block})
        if partial is not None:
            delta_type = "text_delta" if block["type"] == "text" else "input_json_delta"
            key = "text" if delta_type == "text_delta" else "partial_json"
            events.append({"type": "content_block_delta", "index": i, "delta": {"type": delta_type, key: partial}})
        events.append({"type": "content_block_stop", "index": i})
    events.append({"type": "message_delta", "delta": {"stop_reason": stop_reason, "stop_sequence": None},
                   "usage": {"output_tokens": 50, "server_tool_use": {"web_search_requests": 2,
                                                                     "web_fetch_requests": 1}}})
    events.append({"type": "message_stop"})
    return sse(events)


@pytest.fixture
def wired(cfg, monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    requests, responses = [], []

    def handler(request):
        body = json.loads(request.content)
        body["_beta"] = request.headers.get("anthropic-beta", "")
        requests.append(body)
        return httpx2.Response(200, headers={"content-type": "text/event-stream"}, content=responses.pop(0))

    agent = ClaudeAgent(cfg)
    agent.client = anthropic.Anthropic(api_key="test-key", base_url="http://test",
                                       http_client=anthropic.DefaultHttpxClient(transport=httpx2.MockTransport(handler)))
    return agent, requests, responses


def test_research_pause_turn_and_evidence(wired):
    agent, requests, responses = wired
    fetch_use = {"type": "server_tool_use", "id": "srvtoolu_1", "name": "web_fetch", "input": {}}
    fetch_result = {"type": "web_fetch_tool_result", "tool_use_id": "srvtoolu_1", "content": {
        "type": "web_fetch_result", "url": "https://hills.com.au/contact", "content": {
            "type": "document", "source": {"type": "text", "media_type": "text/plain", "data": "info@hills.com.au"}}}}
    record = {"type": "tool_use", "id": "toolu_1", "name": "record_companies", "input": {}}
    responses.append(message([(fetch_use, '{"url": "https://hills.com.au/contact"}'), (fetch_result, None)],
                             "pause_turn"))
    responses.append(message([(record, json.dumps({"companies": [{"company_name": "Hills"}]}))], "tool_use"))

    result = agent.research(system="sys", prompt="find roasters", record_tool=RECORD_COMPANIES, effort="medium",
                            web_searches=3, web_fetches=2)
    assert result.data == {"companies": [{"company_name": "Hills"}]}
    assert result.pages == [("https://hills.com.au/contact", "info@hills.com.au")]

    first = requests[0]
    assert first["model"] == "claude-opus-5-5"
    assert first["fallbacks"] == "default"
    assert "server-side-fallback-2026-07-01" in first["_beta"]
    assert first["output_config"] == {"effort": "medium"}
    assert first["thinking"] == {"type": "adaptive"}
    assert [t.get("type", t.get("name")) for t in first["tools"]] == [
        "web_search_20260209", "web_fetch_20260209", "record_companies"]
    assert first["tools"][2]["strict"] is True
    # After pause_turn the paused assistant turn is re-sent, with no extra user message.
    assert [m["role"] for m in requests[1]["messages"]] == ["user", "assistant"]
    assert agent.usage["web_search_requests"] == 4 and agent.usage["requests"] == 2


def test_research_nudges_when_tool_not_called(wired):
    agent, requests, responses = wired
    responses.append(message([({"type": "text", "text": ""}, "I found two roasters.")], "end_turn"))
    responses.append(message([({"type": "tool_use", "id": "toolu_2", "name": "record_companies", "input": {}},
                               '{"companies": []}')], "tool_use"))
    result = agent.research(system="s", prompt="p", record_tool=RECORD_COMPANIES, effort="low",
                            web_searches=1, web_fetches=0)
    assert result.data == {"companies": []}
    assert [m["role"] for m in requests[1]["messages"]] == ["user", "assistant", "user"]


def test_structured_json(wired):
    agent, requests, responses = wired
    responses.append(message([({"type": "text", "text": ""}, '{"subject": "Hi", "body": "Hello"}')], "end_turn"))
    assert agent.structured(system="s", prompt="p", schema=EMAIL_SCHEMA, effort="medium") == {
        "subject": "Hi", "body": "Hello"}
    assert requests[0]["output_config"]["format"]["type"] == "json_schema"
    assert "tools" not in requests[0]
