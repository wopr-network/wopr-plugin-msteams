# wopr-plugin-msteams

Microsoft Teams channel plugin for WOPR using the Azure Bot Framework.

## Commands

```bash
npm run build     # tsc
npm run check     # biome check + tsc --noEmit (run before committing)
npm run lint:fix  # biome check --fix src/
npm run format    # biome format --write src/
npm test          # vitest run
```

## Architecture

```
src/
  index.ts   # Plugin entry — Bot Framework adapter, Teams activity routing
  types.ts   # Plugin-local types
```

## Key Details

- **Framework**: `botbuilder` (Microsoft Bot Framework SDK v4)
- Requires an Azure Bot registration + App ID/Password
- Teams sends activity payloads to a public HTTPS endpoint — needs Tailscale Funnel or similar for local dev
- Implements `ChannelProvider` from `@wopr-network/plugin-types`
- **Gotcha**: Teams requires proper Bot Framework Manifest in Teams App Studio to enable — not plug-and-play like Discord

## Plugin Contract

Imports only from `@wopr-network/plugin-types`. Never import from `@wopr-network/wopr` core.

## Issue Tracking

All issues in **Linear** (team: WOPR). Issue descriptions start with `**Repo:** wopr-network/wopr-plugin-msteams`.
