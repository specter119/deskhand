"""The installed entry point speaks MCP on stdout and nothing else."""

from __future__ import annotations

import json
import subprocess
import sys

from deskhand import __version__

PYTHON = [sys.executable, "-m", "deskhand"]


def test_version_and_usage():
    version = subprocess.run([*PYTHON, "--version"], capture_output=True, text=True, check=True)
    assert version.stdout.strip() == __version__
    bad = subprocess.run([*PYTHON, "nope"], capture_output=True, text=True)
    assert bad.returncode == 2 and "Unknown app: nope" in bad.stderr and bad.stdout == ""


def test_the_server_answers_json_rpc_on_stdout_only():
    requests = [
        {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2025-06-18",
                "capabilities": {},
                "clientInfo": {"name": "smoke", "version": "0"},
            },
        },
        {"jsonrpc": "2.0", "method": "notifications/initialized"},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
        {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {"name": "append_lines", "arguments": {}},
        },
    ]
    stdin = "".join(json.dumps(request) + "\n" for request in requests)
    server = subprocess.Popen(
        [*PYTHON, "onenote"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
    )
    try:
        assert server.stdin and server.stdout
        server.stdin.write(stdin)
        server.stdin.flush()
        replies = {}
        # Closing stdin ends the session, so keep it open until the last reply has arrived.
        while 3 not in replies:
            line = server.stdout.readline()
            assert line, server.stderr.read() if server.stderr else ""
            reply = json.loads(line)
            replies[reply["id"]] = reply
        server.stdin.close()
        assert server.wait(timeout=30) == 0
    finally:
        server.kill()
        server.wait()
        for stream in (server.stdout, server.stderr):
            if stream:
                stream.close()

    assert replies[1]["result"]["serverInfo"] == {"name": "onenote", "version": __version__}
    assert replies[1]["result"]["instructions"]
    tools = replies[2]["result"]["tools"]
    assert len(tools) == 10
    assert tools[0]["inputSchema"]["type"] == "object" and "outputSchema" in tools[0]

    refused = replies[3]["result"]
    assert refused["isError"] is True
    assert refused["content"][0]["type"] == "text"
    assert refused["content"][0]["text"].startswith("Input validation error")
