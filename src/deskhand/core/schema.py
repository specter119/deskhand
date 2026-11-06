"""A few functions that build the JSON Schema documents the tools advertise.

They emit the same documents the TypeScript implementation did (TypeBox output), because clients cache
and compare them: optional members drop out of `required`, descriptions sit next to the type.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

Schema = dict[str, Any]


@dataclass(frozen=True)
class _Optional:
    schema: Schema


def optional(schema: Schema) -> _Optional:
    return _Optional(schema)


def _described(schema: Schema, description: str | None) -> Schema:
    if description is not None:
        schema["description"] = description
    return schema


def string(description: str | None = None) -> Schema:
    return _described({"type": "string"}, description)


def number(description: str | None = None) -> Schema:
    return _described({"type": "number"}, description)


def boolean(description: str | None = None) -> Schema:
    return _described({"type": "boolean"}, description)


def null() -> Schema:
    return {"type": "null"}


def literal(value: str) -> Schema:
    return {"type": "string", "const": value}


def union(*members: Schema) -> Schema:
    return {"anyOf": list(members)}


def anything() -> Schema:
    return {}


def array(items: Schema, description: str | None = None) -> Schema:
    return _described({"type": "array", "items": items}, description)


def obj(
    properties: dict[str, Schema | _Optional],
    *,
    additional: bool | None = None,
    description: str | None = None,
) -> Schema:
    props: dict[str, Schema] = {}
    required: list[str] = []
    for name, member in properties.items():
        if isinstance(member, _Optional):
            props[name] = member.schema
        else:
            props[name] = member
            required.append(name)
    schema: Schema = {"type": "object", "properties": props}
    if required:
        schema["required"] = required
    if additional is not None:
        schema["additionalProperties"] = additional
    return _described(schema, description)
