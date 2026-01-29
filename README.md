# wopr-plugin-msteams

Microsoft Teams integration for [WOPR](https://github.com/TSavo/wopr) using [Azure Bot Framework](https://docs.microsoft.com/en-us/azure/bot-service/).

## Features

- 💼 **Azure Bot Framework** - Official Microsoft SDK
- 👥 **Channel Support** - Teams, channels, group chats, DMs
- 🧵 **Threading** - Reply in threads or top-level
- 🔒 **DM Policies** - Control access per conversation type
- 👀 **Mention-gated** - Responds to @mentions in channels
- 💬 **Rich Text** - Markdown support
- 📎 **Media** - File attachments and images

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
    
    # Optional
    webhookPort: 3978              # Port for webhook server
    webhookPath: "/api/messages"   # Webhook endpoint path
    requireMention: true           # Require @mention in channels
    replyStyle: "thread"           # "thread" or "top-level"
    dmPolicy: "pairing"            # DM handling policy
    allowFrom: []                  # Allowed user IDs
    groupPolicy: "allowlist"       # Channel/group handling
```

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

## Channel Behavior

### Direct Messages
- Based on `dmPolicy` setting
- `pairing` - All DMs allowed
- `allowlist` - Only allowed users
- `open` - Anyone can DM
- `disabled` - DMs ignored

### Team Channels
- **Mention-gated**: Bot only responds to @mentions (if `requireMention: true`)
- **Reply style**: Can reply in thread or top-level
- **Group policy**: Control which users can trigger in channels

## Troubleshooting

### Bot not responding
1. Check Azure Bot messaging endpoint is correct
2. Verify app ID and password are correct
3. Check Teams app is installed/sideloaded
4. Look at WOPR logs: `wopr logs --follow`

### Webhook errors
- Must use HTTPS in production
- Endpoint must be publicly accessible
- Check firewall/proxy settings

### Authentication errors
- Verify App ID matches Azure Bot registration
- Regenerate client secret if expired
- Check tenant ID is correct

## Security

- ✅ Azure Bot Framework handles authentication
- ✅ Credentials via config or env vars
- ✅ No message content logged
- ✅ HTTPS required for webhooks

## Limitations

- Requires public HTTPS endpoint (no built-in polling)
- Complex Azure/Teams setup compared to other channels
- Webhook server needs to be running continuously

## License

MIT

## See Also

- [WOPR](https://github.com/TSavo/wopr) - The main WOPR project
- [Azure Bot Service](https://docs.microsoft.com/en-us/azure/bot-service/)
- [Teams Bot Documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/what-are-bots)
