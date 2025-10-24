# Production Transformation Summary

This document highlights the intentional divergences between the upstream
[m365-mcp-server](https://github.com/nikolanovoselec/m365-mcp-server) project and
this hardened production fork. Each change notes the motivation, the Cloudflare
services it leverages, and a direct link to the implementation in this
repository.

## Source Code Adjustments

| Area | Description | Reasoning | Cloudflare components | Example lines |
| --- | --- | --- | --- | --- |
| Worker environment bindings | Removed local-only `WORKER_DOMAIN`/`PROTOCOL` flags and introduced the `AI` binding plus optional Access headers so the Worker operates solely on secret-managed configuration. | Ensure production deploys rely on Cloudflare-managed secrets and make AI Gateway + Access context available to downstream logic. | AI Gateway, Access, Durable Objects | [`src/index.ts#L30-L63`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/src/index.ts#L30-L63) |
| Graph client transport | Added `GatewayMetadata`, cached gateway log IDs, and replaced direct `fetch` calls with `env.AI.run("dynamic/microsoft-graph-handler", …)` so every Microsoft Graph request traverses the AI Gateway. | Route all Graph egress through governed Cloudflare infrastructure, attach audit metadata, and surface the gateway log identifier for incident response. | AI Gateway | [`src/microsoft-graph.ts#L71-L639`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/src/microsoft-graph.ts#L71-L639) |
| Durable Object metadata + logging | Before each tool call the agent now builds metadata (user, Access email, Microsoft principal), invokes the Graph client with it, and logs the returned `aiGatewayLogId` for correlation. | Provide end-to-end traceability between MCP tool executions, Access identities, and AI Gateway telemetry without exposing raw tokens. | Access, AI Gateway, Durable Objects | [`src/microsoft-mcp-agent.ts#L109-L218`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/src/microsoft-mcp-agent.ts#L109-L218) |
| Worker configuration | `wrangler.toml` and `.dev.vars` were rewritten with placeholders and explicit secret checklists, while `[[ai]]` bindings became mandatory. | Prevent accidental leakage of tenant-specific IDs and guide operators toward Cloudflare secret storage and AI Gateway configuration. | AI Gateway, Workers KV | [`wrangler.toml#L1-L41`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/wrangler.toml#L1-L41), [`.dev.vars`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/.dev.vars) |

## Documentation Realignment

- `README.md`, `OPERATIONS.md`, and `TECHNICAL.md` emphasise production hardening,
  Cloudflare Access perimeters, AI Gateway routing, and log correlation for
  regulated environments. (See
  [`README.md#L22-L35`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/README.md#L22-L35),
  [`OPERATIONS.md#L134-L141`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/OPERATIONS.md#L134-L141),
  [`TECHNICAL.md#L136-L150`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/TECHNICAL.md#L136-L150)).
- `CONTRIBUTING.md` now directs feature work back to the upstream repository and
  restricts this repo to production/security changes only
  ([`CONTRIBUTING.md#L1-L55`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/CONTRIBUTING.md#L1-L55)).

Together these differences convert the experimental Worker into a deployable,
auditable service that sits behind Cloudflare Access and channels every Microsoft
Graph interaction through Cloudflare AI Gateway.
