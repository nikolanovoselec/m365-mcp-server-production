# Technical Reference – Production Transformation

This document complements the upstream
[m365-mcp-server technical guide](https://github.com/nikolanovoselec/m365-mcp-server/blob/main/TECHNICAL.md)
by describing the additional components required to harden the worker
for enterprise use. It focuses on ingress protection, AI Gateway egress,
and the environment contract that binds everything together.

## 1. Architectural Overview

```mermaid
graph TB
    Client[MCP Client] --> Entry[Worker Entry Point]
    Entry --> Access[Cloudflare Access]
    Access --> OAuth[OAuth Provider]
    OAuth --> SSE["/sse Endpoint"]
    SSE --> DO1[Discovery Durable Object]
    SSE --> DO2[Authenticated Durable Object]
    DO2 --> Gateway[AI Gateway]
    Gateway --> Graph[Microsoft Graph]
    Access -->|CF-Access headers| DO2

    OAuth --> KV1[OAUTH_KV]
    DO2 --> KV2[CONFIG_KV]
```

*Change vs upstream:* the original worker exposed `/sse` directly and called Microsoft Graph with inline
`fetch` statements. Production inserts Cloudflare Access ahead of the Worker and hands every outbound call
to AI Gateway for policy enforcement and auditing. Upstream reference:
https://github.com/nikolanovoselec/m365-mcp-server/blob/main/TECHNICAL.md#system-overview

### Access + OAuth Flow

```mermaid
sequenceDiagram
    participant C as MCP Client
    participant Access as Cloudflare Access
    participant W as Worker
    participant O as OAuth Provider
    participant MS as Microsoft Entra

    C->>Access: 1. Request /sse
    Access-->>C: 2. Enforce SSO / MFA
    C->>Access: 3. Present Access session
    Access->>W: 4. Forward request + CF-Access headers
    W->>O: 5. /authorize (approval + state)
    O-->>C: 6. Approval dialog (first visit)
    C->>O: 7. Approve + redirect
    O->>MS: 8. Authorize user
    MS-->>O: 9. Authorization code
    O->>MS: 10. Token exchange
    MS-->>O: 11. Access & refresh tokens
    O->>W: 12. Persist tokens in props
    W-->>C: 13. OAuth access token for MCP session
```

- Change vs upstream: the original flow skipped the Access perimeter and stored tokens immediately. Production
  requires the Access cookie before hitting `/authorize` and logs Access headers alongside Microsoft tokens.
  See upstream diagram: https://github.com/nikolanovoselec/m365-mcp-server/blob/main/TECHNICAL.md#oauth-flow-architecture

### AI Gateway Egress

```mermaid
sequenceDiagram
    participant DO as Durable Object
    participant W as Worker
    participant G as AI Gateway
    participant Graph as Microsoft Graph

    DO->>W: 1. Tool call + Microsoft tokens
    W->>G: 2. env.AI.run(dynamic route, metadata)
    G->>Graph: 3. Forward request / policy controls
    Graph-->>G: 4. API response
    G-->>W: 5. Response + `aiGatewayLogId`
    W->>DO: 6. Tool result (logs metadata + log ID)
```

- Change vs upstream: calls now route through AI Gateway instead of direct `fetch`, enabling policy enforcement
  and log correlation. Upstream sequence: https://github.com/nikolanovoselec/m365-mcp-server/blob/main/TECHNICAL.md#microsoft-graph-integration

1. **Cloudflare Access** acts as checkpoint #1 (perimeter). Requests without a valid Access token never
   reach the Worker. Identity, device posture, and service token claims can be surfaced through headers
   (`CF-Access-Authenticated-User-Email`, `CF-Access-Jwt-Assertion`).
2. **Worker & Durable Object** preserve the MCP protocol behaviours established upstream, including
   Microsoft OAuth 2.1 storage and dual-token props.
3. **AI Gateway** intercepts every outbound call. Dynamic routes encapsulate corporate policies: logging,
   caching, DLP, and rate limiting. Metadata from the Worker identifies the user and MCP tool invoked.

- Diagram above replaces the upstream “direct fetch” flow: Access now guards `/sse`, and every outbound channel is
  routed through AI Gateway before returning to the Durable Object.

## 2. Environment Contract

```ts
export interface Env {
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_TENANT_ID: string;
  GRAPH_API_VERSION: string;
  MICROSOFT_CLIENT_SECRET: string;
  ENCRYPTION_KEY: string;
  COOKIE_ENCRYPTION_KEY: string;
  COOKIE_SECRET: string;

  MCP_OBJECT: DurableObjectNamespace;
  OAUTH_KV: KVNamespace;

  /** Cloudflare AI Gateway binding */
  AI: Ai;

  /** Optional Access headers (surfaces identity + posture) */
  CF_Access_Jwt_Assertion?: string;
  CF_Access_Authenticated_User_Email?: string;
  CF_Access_Authenticated_User_Id?: string;
}
```

**Key differences vs upstream:**
- Secrets are not expressed in `[vars]`; they must be supplied via `wrangler secret`.
- The `AI` binding is mandatory when deploying to the hardened environment.
- Access headers are optional, but logging them enables correlation between SSO identities
  and Microsoft OAuth users.

## 3. AI Gateway Invocation Pattern

### Dynamic Route Definition

Create a dynamic route named `dynamic/microsoft-graph-handler` in the AI Gateway UI. Configure it to
forward requests to `https://graph.microsoft.com`. Additional routes (e.g., `dynamic/llm-summarizer`)
can proxy other downstream systems.

### Worker Usage

```ts
const response = await env.AI.run(
  "dynamic/microsoft-graph-handler",
  {
    method: "POST",
    path: "/v1.0/me/sendMail",
    headers: {
      Authorization: `Bearer ${microsoftAccessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  },
  {
    gateway: {
      id: "m365-egress-gateway",
      metadata: {
        userId: this.props?.id ?? "unknown",
        mcpTool: "sendEmail",
        requestId,
        userEmail: env.CF_Access_Authenticated_User_Email ?? this.props?.mail,
      },
    },
  },
);
```

- **`path`**: Relative to Microsoft Graph base URL, allowing the gateway to centralise origin logic.
- **Metadata**: Supply user identifier, MCP tool name, and correlation IDs so that gateway logs
  support incident response and analytics.
- **Gateway helpers**: Capture `env.AI.aiGatewayLogId` for the most recent call or invoke
  `env.AI.gateway("m365-egress-gateway").patchLog(...)` / `getLog(...)` when you need to append
  metadata or fetch request bodies ([binding methods](https://developers.cloudflare.com/ai-gateway/integrations/worker-binding-methods/)).
- **Error Handling**: The worker should translate non-2xx responses into structured MCP errors,
  indicating whether the issue is policy (429/DLP) vs Graph-specific (403/401).

## 4. Access Awareness

Cloudflare Access for SaaS issues the OAuth token that the worker presents on each request. If the worker later needs to call internal HTTP applications, configure `linked_app_token` policies so the same token is honoured downstream ([docs](https://developers.cloudflare.com/cloudflare-one/access-controls/applications/http-apps/mcp-servers/linked-apps/)).

The worker can read Access-derived headers to enrich logs or enforce additional checks, for example:

```ts
const userEmail = env.CF_Access_Authenticated_User_Email;
if (userEmail) {
  console.log(`Access-authenticated user: ${userEmail}`);
}
```

This data can be injected into the AI Gateway metadata payload for end-to-end traceability.

## 5. Durable Object Considerations

- Discovery mode continues to run without OAuth props; Access ensures only authorised clients
  reach this stage.
- Authenticated tool executions inherit both Access context (perimeter) and Microsoft OAuth tokens
  (application). Ensure Durable Object logs do not emit sensitive tokens; rely on metadata instead.

## 6. Logging & Observability

- **AI Gateway**: Primary location for monitoring outbound traffic, rate limiting, and DLP violations.
  - Dynamic routes expose provider/model decisions and quotas ([docs](https://developers.cloudflare.com/ai-gateway/features/dynamic-routing/)).
- **Gateway Log Correlation**: Every Graph request attaches metadata (`userId`, `userEmail`, `mcpTool`, `requestId`) and records the resulting `env.AI.aiGatewayLogId`. Durable Object logs emit this identifier so operators can jump directly to the corresponding entry inside the AI Gateway dashboard.
- **Workers Tail**: Use `wrangler tail --metadata` to surface request IDs and Access identity info.
- **Access Audit Logs**: Provide authentication history, device posture evaluation, and policy results.
- **Microsoft Entra ID**: Audit application sign-ins to confirm OAuth flows remain compliant.

## 7. Security Checklist

1. Access required and MFA enforced before reaching `/sse`.
2. Secrets only exist within Cloudflare secret storage.
3. AI Gateway metadata consistently labels requests (`userId`, `mcpTool`, `requestId`, `userEmail` when available).
4. Logs contain no raw OAuth tokens or Microsoft responses beyond what is necessary.
5. Deployment scripts run `npm run validate` to preserve lint/format/type safety.

---

For core MCP protocol behaviour, tool schemas, and base architecture, continue to reference the upstream
[m365-mcp-server documentation](https://github.com/nikolanovoselec/m365-mcp-server/blob/main/TECHNICAL.md).
