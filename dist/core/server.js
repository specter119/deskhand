import { fromJsonSchema } from "@modelcontextprotocol/server";
/** TypeBox emits plain JSON Schema, which the SDK accepts once it is wrapped. */
export function jsonSchema(schema) {
    return fromJsonSchema(schema);
}
/** A thrown error becomes an isError tool result, so `run` just throws on failure. */
export function defineTool(server, tool) {
    server.registerTool(tool.name, {
        title: tool.title,
        description: tool.description,
        inputSchema: jsonSchema(tool.input),
        outputSchema: jsonSchema(tool.output),
        annotations: tool.annotations,
    }, async (args, ctx) => {
        const result = await tool.run(args, { signal: ctx.mcpReq.signal });
        return {
            content: [{ type: "text", text: result.text }],
            structuredContent: result.structured,
        };
    });
}
