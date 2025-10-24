# Production Transformation Guide for Microsoft 365 MCP Server

This repository captures the hardened deployment of the open-source [m365-mcp-server](https://github.com/nikolanovoselec/m365-mcp-server) project.
The objective is to evolve the experimental Cloudflare Worker into an enterprise-grade agent with
clear ingress controls, governed egress, and auditable operations.

> Looking for feature-level behaviour, tool definitions, or local development instructions?
> See the upstream repository documentation:
> - [README](https://github.com/nikolanovoselec/m365-mcp-server/blob/main/README.md)
> - [OPERATIONS](https://github.com/nikolanovoselec/m365-mcp-server/blob/main/OPERATIONS.md)
> - [TECHNICAL](https://github.com/nikolanovoselec/m365-mcp-server/blob/main/TECHNICAL.md)

## Scope

- **Source code** mirrors the upstream worker while introducing AI Gateway bindings,
  hardened environment typings, and security-centric defaults.
- **Documentation** in this repository focuses exclusively on production-readiness
  and the transformation steps required to land inside a secured Cloudflare estate.
- **Operations** material provides deterministic procedures for deployment,
  validation, logging, and change management.

## Security Transformation Goals

1. **Perimeter Enforcement** – Protect the worker with Cloudflare Access
   (SSO, MFA, device posture, service tokens) before any MCP handshake occurs.
2. **Application Authorisation** – Preserve Microsoft OAuth 2.1 + PKCE flows while
   storing secrets exclusively through `wrangler secret`.
3. **Egress Governance** – Route all outbound API calls (Microsoft Graph, LLMs, webhooks)
   through Cloudflare AI Gateway with metadata for logging and policy enforcement.
4. **Operational Guardrails** – Maintain reproducible deployment, observability, and
  incident response checklists suitable for regulated environments.
5. **Auditable Graph Egress** – All Microsoft Graph calls run through Cloudflare AI
  Gateway with enriched metadata (`userId`, `userEmail`, `mcpTool`, `requestId`) and
  expose the gateway log identifier so investigations can pivot between Workers logs
  and gateway analytics.

## How This Fork Differs from Upstream

| Area | Description | Reasoning | Cloudflare components | Example lines |
| --- | --- | --- | --- | --- |
| Worker environment bindings | Drops local testing flags and adds the `AI` binding plus optional Cloudflare Access headers so runtime configuration is supplied exclusively via secrets. | Keep production deployments secret-driven and expose Access context to downstream handlers. | AI Gateway, Access, Durable Objects | [`src/index.ts#L30-L63`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/src/index.ts#L30-L63) |
| Graph client transport | Introduces `GatewayMetadata`, proxies requests through `env.AI.run("dynamic/microsoft-graph-handler", …)`, and captures `aiGatewayLogId` after each call. | Ensure every Microsoft Graph request traverses Cloudflare AI Gateway with audit metadata available for incident response. | AI Gateway | [`src/microsoft-graph.ts#L66-L639`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/src/microsoft-graph.ts#L66-L639) |
| Durable Object metadata + logging | Builds per-tool metadata from Access/Microsoft identities, forwards it to the Graph client, and logs the resulting gateway correlation ID. | Provide traceability between MCP tool executions, Access identities, and AI Gateway analytics without exposing raw tokens. | Access, AI Gateway, Durable Objects | [`src/microsoft-mcp-agent.ts#L109-L218`](https://github.com/nikolanovoselec/m365-mcp-server-production/blob/main/src/microsoft-mcp-agent.ts#L109-L218) |
| Worker configuration | `wrangler.example.toml` provides placeholder bindings, mandatory `[[ai]]` configuration, and explicit instructions for Cloudflare secrets. | Prevent accidental disclosure of tenant-specific identifiers and guide operators toward Cloudflare secret storage. | AI Gateway, Workers KV | [`wrangler.example.toml`](wrangler.example.toml) |

## Required Cloudflare Secrets

| Secret | Purpose | Required | Provisioning |
| --- | --- | --- | --- |
| `MICROSOFT_CLIENT_ID` | Microsoft Entra application (client) ID used during OAuth flows. | ✅ | `wrangler secret put MICROSOFT_CLIENT_ID` |
| `MICROSOFT_TENANT_ID` | Microsoft Entra tenant identifier for token endpoints. | ✅ | `wrangler secret put MICROSOFT_TENANT_ID` |
| `GRAPH_API_VERSION` | Microsoft Graph API version (e.g., `v1.0`) used when constructing routes. | ✅ | `wrangler secret put GRAPH_API_VERSION` |
| `MICROSOFT_CLIENT_SECRET` | Microsoft Entra application secret used for code/refresh exchanges. Must never live in `wrangler.toml` or environment variables. | ✅ | `wrangler secret put MICROSOFT_CLIENT_SECRET` |
| `ENCRYPTION_KEY` | 32-byte hex key for encrypting Durable Object state and OAuth props. | ✅ | `wrangler secret put ENCRYPTION_KEY` |
| `COOKIE_ENCRYPTION_KEY` | 32-byte hex key for Access approval cookie encryption. | ✅ | `wrangler secret put COOKIE_ENCRYPTION_KEY` |
| `COOKIE_SECRET` | HMAC secret for signing Access approval cookies. | ✅ | `wrangler secret put COOKIE_SECRET` |
| `AI_GATEWAY_SERVICE_TOKEN` | Service token for AI Gateway if policies require authenticated Worker access. | Optional | `wrangler secret put AI_GATEWAY_SERVICE_TOKEN` |

## Repository Map

- [`OPERATIONS.md`](./OPERATIONS.md) – Phase-by-phase migration playbook, secret strategy,
  validation checklist, and maintenance tasks.
- [`TECHNICAL.md`](./TECHNICAL.md) – Architectural deep dive covering Access, Durable Objects,
  AI Gateway invocation patterns, environment contract, and logging strategy.
- `src/` – Worker source aligned with the hardened environment.
- `wrangler.toml` – Production configuration (AI binding, Durable Object, Access-friendly routes).

## Transformation Summary

Around 260 lines of worker code (spread across `src/index.ts`, `src/microsoft-graph.ts`, and
`src/microsoft-mcp-agent.ts`) were adjusted, alongside a handful of Cloudflare console changes, to
turn the open prototype into an enterprise-hardened Microsoft 365 Remote MCP Server. Security controls added:

- **Cloudflare Access perimeter** – `/sse` now sits behind SSO/MFA/device posture, and Access headers reach the Durable Object for auditing.
- **Cloudflare AI Gateway egress** – Every Graph call flows through dynamic routes with policy enforcement, logging, and DLP.
- **Secret lifecycle via wrangler** – No credentials in `[vars]`; secrets are stored/rotated with `wrangler secret put`.
- **Durable Object correlation** – Tool executions log the Cloudflare `aiGatewayLogId`, linking MCP activity to Gateway telemetry.
- **Configuration hygiene** – Example configs rely on placeholders and current compatibility dates; production `wrangler.toml` consumes secrets exclusively.

## Migration Roadmap

1. **Provision Cloudflare Access** – create a self-hosted application guarding `mcp.<domain>`
   and authorise relevant identities (human + automation).
2. **Deploy AI Gateway** – create `m365-egress-gateway`, enable logging/rate limiting/DLP,
   and define the required dynamic routes (e.g., `dynamic/microsoft-graph-handler`).
3. **Refactor Configuration & Secrets** – bind the gateway via `[[ai]]`, remove `[vars]`,
  populate the placeholder IDs in `wrangler.toml`, and push all credentials using
  `wrangler secret put`.
4. **Update Source Code** – replace direct `fetch` invocations with `env.AI.run(...)`,
   attach `gateway.metadata` (user identifiers, MCP tool name, correlation IDs),
   and centralise error handling for Access or gateway denials.
5. **Deploy & Validate** – run `wrangler deploy --env production`, complete the Access SSO,
   Microsoft consent, and tool execution flow; confirm AI Gateway/Access telemetry.

### Transformation at a Glance

```mermaid
flowchart LR
    A[Early MCP Worker<br>Direct fetch + inline secrets] --> B[Phase 1<br>Cloudflare Access App]
    B --> C[Phase 2<br>MCP Portal + Linked Apps]
    C --> D[Phase 3<br>AI Gateway Dynamic Routes]
    D --> E[Phase 4<br>Code Refactor<br>env.AI.run + metadata]
    E --> F[Phase 5<br>Deploy & Validate<br>Access SSO • Gateway logs]
```
- Change vs upstream: Access now fronts `/sse`, AI Gateway owns all Graph egress, and Durable Objects log
  `aiGatewayLogId` values for audit trails. Upstream timeline:
  https://github.com/nikolanovoselec/m365-mcp-server/blob/main/README.md#why-did-i-build-this

## Reference Material

- Cloudflare docs – [AI Gateway binding methods](https://developers.cloudflare.com/ai-gateway/integrations/worker-binding-methods/)
- Cloudflare docs – [Universal endpoint](https://developers.cloudflare.com/ai-gateway/usage/universal/)
- Cloudflare docs – [Dynamic routing](https://developers.cloudflare.com/ai-gateway/features/dynamic-routing/)
- Cloudflare docs – [Access linked apps for MCP servers](https://developers.cloudflare.com/cloudflare-one/access-controls/applications/http-apps/mcp-servers/linked-apps/)
- Cloudflare GitHub – [AI Gateway MCP server](https://github.com/cloudflare/mcp-server-cloudflare/tree/main/apps/ai-gateway) (log tooling & OAuth patterns)
- Cloudflare docs – [Durable Objects](https://developers.cloudflare.com/durable-objects/) (stateful session management)

## Contributing

All feature work, tool enhancements, or general documentation improvements should originate in
[m365-mcp-server](https://github.com/nikolanovoselec/m365-mcp-server). Contributions here should
focus on production-hardening, deployment automation, and operational playbooks. When updating
source code, keep both repositories in sync to ensure transformation guidance remains accurate.
