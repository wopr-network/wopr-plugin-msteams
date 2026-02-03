# wopr-plugin-msteams

[![npm version](https://img.shields.io/npm/v/wopr-plugin-msteams.svg)](https://www.npmjs.com/package/wopr-plugin-msteams)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![WOPR](https://img.shields.io/badge/WOPR-Plugin-blue)](https://github.com/TSavo/wopr)

Microsoft Teams integration for [WOPR](https://github.com/TSavo/wopr) using [Azure Bot Framework](https://docs.microsoft.com/en-us/azure/bot-service/).

> Part of the [WOPR](https://github.com/TSavo/wopr) ecosystem - Self-sovereign AI session management over P2P.

## Features

- 💼 **Azure Bot Framework** - Official Microsoft SDK (botbuilder v4.22+)
- 👥 **Channel Support** - Teams channels, group chats, and direct messages
- 🧵 **Threading** - Configurable reply in threads or top-level
- 🔒 **Access Policies** - Separate DM and group/channel policies with allowlists
- 👀 **Mention-gated** - Optionally require @mentions in channels/groups
- 💬 **Markdown** - Responses support markdown formatting

## Prerequisites

### 1. Create Azure Bot Resource

1. Go to [Azure Portal](https://portal.azure.com)
2. Create a **Azure Bot** resource
3. Note the **Microsoft App ID**
4. Create **Client Secret** (App Password)
5. Note your **Azure AD Tenant ID**

### 2. Configure Messaging Endpoint

In Azure Bot configuration:
- Set **Messaging endpoint** to your public URL + `/api/messages`
- Example: `https://your-server.com/api/messages`

### 3. Create Teams App

1. Go to [Teams Developer Portal](https://dev.teams.microsoft.com/)
2. Create a new app
3. Add **Bot** capability
4. Enter your Azure Bot App ID
5. Download app package and install/sideload

## Installation

```bash
wopr channels add msteams
```

Or manually:
```bash
npm install wopr-plugin-msteams
```

## Configuration

```yaml
# ~/.wopr/config.yaml
channels:
  msteams:
    # Required - Azure Bot credentials
    appId: "00000000-0000-0000-0000-000000000000"
    appPassword: "your-client-secret"
    tenantId: "00000000-0000-0000-0000-000000000000"

    # Optional - Webhook settings
    webhookPort: 3978              # Port for webhook server (default: 3978)
    webhookPath: "/api/messages"   # Webhook endpoint path (default: /api/messages)

    # Optional - Channel behavior
    requireMention: true           # Require @mention in channels/groups (default: true)
    replyStyle: "thread"           # "thread" or "top-level" (default: thread)

    # Optional - DM access control
    dmPolicy: "pairing"            # "pairing" | "allowlist" | "open" | "disabled"
    allowFrom: []                  # User IDs for DM allowlist

    # Optional - Group/channel access control
    groupPolicy: "allowlist"       # "allowlist" | "open" | "disabled"
    groupAllowFrom: []             # User IDs for group allowlist (falls back to allowFrom)
```

### Policy Options

| Policy | Behavior |
|--------|----------|
| `open` | Anyone can message |
| `pairing` | All DMs allowed (DM only) |
| `allowlist` | Only listed user IDs allowed |
| `disabled` | Messages ignored |

**Note:** Use `"*"` in `allowFrom` or `groupAllowFrom` to allow all users.

## Environment Variables

| Variable | Description |
|----------|-------------|
| `MSTEAMS_APP_ID` | Azure Bot App ID |
| `MSTEAMS_APP_PASSWORD` | Client Secret |
| `MSTEAMS_TENANT_ID` | Azure AD Tenant ID |

## Architecture

```
┌─────────────────┐      HTTPS Webhook       ┌─────────────────┐
│   Microsoft     │ ◄──────────────────────► │   WOPR Plugin   │
│   Teams         │      Bot Framework       │  (Azure Bot)    │
└─────────────────┘                          └─────────────────┘
        │                                            │
        │                                            │
┌─────────────────┐                          ┌─────────────────┐
│   Azure Bot     │                          │      WOPR       │
│   Service       │                          │     Core        │
└─────────────────┘                          └─────────────────┘
```

## Webhook Setup

The plugin requires a **public HTTPS endpoint** for Teams to send messages to.

### Development Options:

1. **ngrok** (for local development):
   ```bash
   ngrok http 3978
   # Use the HTTPS URL + /api/messages in Azure Bot config
   ```

2. **Cloudflare Tunnel**:
   ```bash
   cloudflared tunnel --url http://localhost:3978
   ```

3. **Production**: Use your server's public URL with SSL

## Message Flow

### Direct Messages (Personal Chats)
Access controlled by `dmPolicy`:
- `pairing` (default) - All DMs are processed
- `allowlist` - Only users in `allowFrom` list
- `open` - Anyone can DM
- `disabled` - DMs are ignored

### Team Channels and Group Chats
Access controlled by `groupPolicy` and `requireMention`:
- **Mention requirement**: When `requireMention: true` (default), the bot only responds when @mentioned
- **Access control**: Uses `groupPolicy` with `groupAllowFrom` (falls back to `allowFrom`)
- **Reply style**: `thread` replies to the original message, `top-level` posts as a new message

### Session Keys
Each conversation gets a unique session key: `msteams-{conversationId}`

## Troubleshooting

### Bot not responding
1. Check Azure Bot messaging endpoint is correct
2. Verify app ID, password, and tenant ID are correct
3. Check Teams app is installed/sideloaded
4. Look at plugin logs: `~/.wopr/logs/msteams-plugin.log`
5. Check error logs: `~/.wopr/logs/msteams-plugin-error.log`
6. In channels, ensure you're @mentioning the bot (if `requireMention: true`)

### Webhook errors
- Must use HTTPS in production
- Endpoint must be publicly accessible
- Check firewall/proxy settings
- Verify the webhook handler is properly integrated

### Authentication errors
- Verify App ID matches Azure Bot registration
- Regenerate client secret if expired
- Check tenant ID is correct
- All three credentials (appId, appPassword, tenantId) are required

### Policy blocking messages
- Check `dmPolicy` for direct messages
- Check `groupPolicy` and `requireMention` for channels
- Verify user IDs in `allowFrom`/`groupAllowFrom` lists

## Programmatic Usage

The plugin exports a webhook handler for integration with your HTTP server:

```typescript
import plugin, { handleWebhook } from "wopr-plugin-msteams";

// Initialize the plugin with WOPR context
await plugin.init(woprContext);

// In your Express/Fastify/etc. server:
app.post("/api/messages", async (req, res) => {
  await handleWebhook(req, res);
});
```

## Security

- Azure Bot Framework handles authentication via `ConfigurationBotFrameworkAuthentication`
- Credentials via config file or environment variables
- Message metadata logged, content passed to WOPR
- HTTPS required for webhooks in production
- Built-in error handling with user-friendly error messages

## Limitations

- Requires public HTTPS endpoint (no built-in polling mode)
- Complex Azure/Teams setup compared to other channels
- Webhook server must be running continuously
- No adaptive card support yet (text/markdown only)
- No file attachment handling yet

## Dependencies

- `botbuilder` ^4.22.0 - Microsoft Bot Framework SDK
- `winston` ^3.11.0 - Logging

## License

MIT

## See Also

- [WOPR](https://github.com/TSavo/wopr) - The main WOPR project
- [Azure Bot Service](https://docs.microsoft.com/en-us/azure/bot-service/)
- [Teams Bot Documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/what-are-bots)
