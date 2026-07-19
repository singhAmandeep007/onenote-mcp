/**
 * An MCP stdio server communicates with the client over stdout using JSON-RPC, so
 * anything written to stdout that is not a protocol message will corrupt the stream.
 * All diagnostic logging therefore goes to stderr.
 */
export function log(...args: unknown[]): void {
  console.error('[onenote-mcp]', ...args);
}
