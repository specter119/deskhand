"""Turns a list of tool definitions into an MCP server.

The SDK does not validate tool arguments or results, so this module does it, in the order the TypeScript SDK
used: missing arguments become an empty object, input is validated, the handler runs, and its structured
result is validated against the output schema. Every failure after tool lookup becomes an `isError` result
that carries only a message. An unknown tool name is a protocol error.
"""

from __future__ import annotations

from collections.abc import AsyncGenerator, Awaitable, Callable
from contextlib import asynccontextmanager
from dataclasses import dataclass
from typing import Any

from jsonschema import Draft202012Validator
from jsonschema.protocols import Validator
from mcp import MCPError
from mcp.server import Server, ServerRequestContext
from mcp.types import (
    INVALID_PARAMS,
    CallToolRequestParams,
    CallToolResult,
    ListToolsResult,
    PaginatedRequestParams,
    TextContent,
    Tool,
    ToolAnnotations,
)

MAX_VALIDATION_MESSAGE = 500


@dataclass(frozen=True)
class ToolOutput:
    text: str
    """What a model reads."""
    structured: dict[str, Any]
    """What a script reads; validated against the tool's output schema."""


@dataclass(frozen=True)
class ToolSpec:
    name: str
    title: str
    description: str
    input_schema: dict[str, Any]
    output_schema: dict[str, Any]
    annotations: ToolAnnotations
    run: Callable[[dict[str, Any]], Awaitable[ToolOutput]]
    """Throws on failure; the message becomes the `isError` text."""


def _first_error(validator: Validator, value: Any) -> str | None:
    error = next(iter(sorted(validator.iter_errors(value), key=lambda e: list(e.absolute_path))), None)
    if error is None:
        return None
    where = "/".join(str(part) for part in error.absolute_path)
    message = f"{where}: {error.message}" if where else error.message
    return message[:MAX_VALIDATION_MESSAGE]


def _failure(
    message: str, *, code: str | None = None, details: dict[str, Any] | None = None
) -> CallToolResult:
    structured = {"error": {"code": code or "error", "details": details}} if details is not None else None
    return CallToolResult(content=[TextContent(text=message)], is_error=True, structured_content=structured)


def output_problem(schema: dict[str, Any], value: Any) -> str | None:
    """The same check `build_server` applies, exposed for a write that must know before it returns."""
    return _first_error(Draft202012Validator(schema), value)


def build_server(
    name: str,
    version: str,
    instructions: str,
    specs: list[ToolSpec],
    on_close: Callable[[], Awaitable[None]] | None = None,
) -> Server:
    """`on_close` runs when the server's run loop ends, for example to stop a helper process."""
    by_name = {spec.name: spec for spec in specs}
    input_validators = {spec.name: Draft202012Validator(spec.input_schema) for spec in specs}
    output_validators = {spec.name: Draft202012Validator(spec.output_schema) for spec in specs}
    tools = [
        Tool(
            name=spec.name,
            title=spec.title,
            description=spec.description,
            input_schema=spec.input_schema,
            output_schema=spec.output_schema,
            annotations=spec.annotations,
        )
        for spec in specs
    ]

    async def list_tools(
        ctx: ServerRequestContext[Any], params: PaginatedRequestParams | None
    ) -> ListToolsResult:
        return ListToolsResult(tools=tools)

    async def call_tool(ctx: ServerRequestContext[Any], params: CallToolRequestParams) -> CallToolResult:
        spec = by_name.get(params.name)
        if spec is None:
            raise MCPError(INVALID_PARAMS, f"Tool {params.name} not found")

        arguments = params.arguments or {}
        problem = _first_error(input_validators[spec.name], arguments)
        if problem:
            return _failure(f"Input validation error: {problem}")

        try:
            output = await spec.run(arguments)
        except Exception as error:
            details = getattr(error, "details", None)
            code = getattr(error, "code", None)
            return _failure(
                str(error) or type(error).__name__,
                code=code if isinstance(code, str) else None,
                details=details if isinstance(details, dict) else None,
            )

        problem = _first_error(output_validators[spec.name], output.structured)
        if problem:
            # A write handler's result can fail validation after the write call returned. Do not let that
            # read as a plain failure with nothing written.
            returned = (
                output.structured.get("status") == "applied" or output.structured.get("applied") is True
            )
            note = " The write call had already returned; read the page before retrying." if returned else ""
            return _failure(f"Output validation error: {problem}{note}")
        return CallToolResult(content=[TextContent(text=output.text)], structured_content=output.structured)

    @asynccontextmanager
    async def lifespan(_: Server[Any]) -> AsyncGenerator[dict[str, Any]]:
        try:
            yield {}
        finally:
            if on_close is not None:
                await on_close()

    return Server(
        name,
        version=version,
        instructions=instructions,
        on_list_tools=list_tools,
        on_call_tool=call_tool,
        lifespan=lifespan,
    )
