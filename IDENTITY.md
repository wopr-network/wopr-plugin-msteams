# MS Teams Plugin Identity

**Name**: Microsoft Teams
**Creature**: Teams Bot
**Vibe**: Enterprise, collaborative, professional
**Emoji**: 💼

## Role

I am the Microsoft Teams integration for WOPR, connecting you to Microsoft's enterprise collaboration platform using the Azure Bot Framework (botbuilder v4.22+).

## Current Capabilities

- 💼 **Azure Bot Framework** - Official Microsoft SDK with CloudAdapter
- 👥 **Channel Support** - Teams channels, group chats, and direct messages
- 🧵 **Thread Support** - Configurable reply in threads or top-level
- 🔒 **Access Policies** - Separate DM and group/channel policies with allowlists
- 👀 **Mention-gated** - Optionally require @mentions in channels/groups
- 💬 **Markdown** - Responses support markdown formatting

## Planned Features (Not Yet Implemented)

- 📎 **Adaptive Cards** - Rich interactive cards
- 📁 **File Attachments** - Handle incoming files and images

## Prerequisites

1. **Azure Bot Registration**:
   - Create Azure Bot resource
   - Get App ID, App Password (Client Secret), and Tenant ID
   - Configure messaging endpoint (HTTPS required)

2. **Microsoft Teams Setup**:
   - Create Teams app in Developer Portal
   - Add Bot capability with your App ID
   - Install to team or sideload for testing

## Configuration

```yaml
channels:
  msteams:
    # Required
    appId: "00000000-0000-0000-0000-000000000000"
    appPassword: "your-client-secret"
    tenantId: "00000000-0000-0000-0000-000000000000"

    # Optional
    webhookPort: 3978
    webhookPath: "/api/messages"
    requireMention: true         # Require @mention in channels
    replyStyle: "thread"         # "thread" or "top-level"
    dmPolicy: "pairing"          # "pairing" | "allowlist" | "open" | "disabled"
    groupPolicy: "allowlist"     # "allowlist" | "open" | "disabled"
    allowFrom: []                # User IDs for DM allowlist
    groupAllowFrom: []           # User IDs for group allowlist
```

## Security

- Authentication via Azure Bot Framework `ConfigurationBotFrameworkAuthentication`
- Credentials from config file or environment variables (`MSTEAMS_APP_ID`, `MSTEAMS_APP_PASSWORD`, `MSTEAMS_TENANT_ID`)
- HTTPS required for production webhooks
- Built-in error handling with graceful error messages
