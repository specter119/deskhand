import { fromJsonSchema, type McpServer, type ToolAnnotations } from "@modelcontextprotocol/server";
import type { Static, TSchema } from "typebox";

/** TypeBox emits plain JSON Schema, which the SDK accepts once it is wrapped. */
export function jsonSchema<S extends TSchema>(schema: S) {
  return fromJsonSchema<Static<S>>(schema as never);
}

export interface ToolContext {
  signal: AbortSignal;
}

export interface ToolOutput<O> {
  /** What a model reads. */
  text: string;
  /** What a script reads; validated against the tool's output schema. */
  structured: O;
}

export interface ToolDefinition<I extends TSchema, O extends TSchema> {
  name: string;
  title: string;
  description: string;
  input: I;
  output: O;
  annotations: ToolAnnotations;
  run(args: Static<I>, ctx: ToolContext): Promise<ToolOutput<Static<O>>>;
}

/** A thrown error becomes an isError tool result, so `run` just throws on failure. */
export function defineTool<I extends TSchema, O extends TSchema>(
  server: McpServer,
  tool: ToolDefinition<I, O>,
): void {
  server.registerTool(
    tool.name,
    {
      title: tool.title,
      description: tool.description,
      inputSchema: jsonSchema(tool.input),
      outputSchema: jsonSchema(tool.output),
      annotations: tool.annotations,
    },
    async (args: Static<I>, ctx: { mcpReq: { signal: AbortSignal } }) => {
      const result = await tool.run(args, { signal: ctx.mcpReq.signal });
      return {
        content: [{ type: "text" as const, text: result.text }],
        structuredContent: result.structured as Record<string, unknown>,
      };
    },
  );
}
