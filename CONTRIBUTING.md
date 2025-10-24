# Contributing – Production Transformation Track

This repository exists to document and maintain the enterprise hardening of
[m365-mcp-server](https://github.com/nikolanovoselec/m365-mcp-server).
All feature development, protocol changes, or Microsoft Graph enhancements should
continue to land in the upstream project. Contributions here must keep both
repositories aligned while focusing on production deployment, security controls,
and operational resilience.

## Contribution Areas

- Documentation covering Access, AI Gateway, or security posture updates
- Production configuration changes (`wrangler.toml`, bindings, secret handling)
- Operational scripts, validation checklists, or monitoring improvements
- Bug fixes specific to the hardened environment (e.g., gateway error translation)

## Workflow

1. Sync the latest changes from upstream `m365-mcp-server`.
2. Apply or adjust production-specific patches in this repository.
3. Run validation locally (`npm run validate`).
4. Document notable adjustments in `OPERATIONS.md` and/or `TECHNICAL.md`.
5. Submit a pull request summarising:
   - Upstream baseline commit (if relevant)
   - Production changes introduced
   - Validation steps performed (Access, OAuth, AI Gateway)

## Local Environment

Use the upstream repository for feature development workflows. When verifying
production changes locally, mock AI Gateway calls or use a dedicated staging
gateway to avoid polluting production telemetry.

```bash
npm install
npm run validate
wrangler deploy --env staging
```

Secrets should never be committed or stored in `.dev.vars` when working inside this
repository. Prefer `wrangler secret` even for staging environments.

## Code & Documentation Standards

- Follow existing TypeScript, ESLint, and Prettier configurations.
- Keep comments concise; favour documentation for detailed explanations.
- Maintain professional tone and ensure docs link back to upstream material when needed.

## Review Expectations

- Pull requests must include evidence of AI Gateway and Access validation.
- Changes affecting security boundaries require explicit testing notes.
- Ensure backwards compatibility with existing Access policies and gateway routes.

Thank you for helping maintain the production-ready posture of the Microsoft 365 MCP Server.
