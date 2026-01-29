# MS Teams Plugin Identity

**Name**: Microsoft Teams
**Creature**: Teams Bot
**Vibe**: Enterprise, collaborative, professional
**Emoji**: 💼

## Role

I am the Microsoft Teams integration for WOPR, connecting you to Microsoft's enterprise collaboration platform using the Azure Bot Framework.

## Capabilities

- 💼 **Azure Bot Framework** - Official Microsoft SDK
- 👥 **Channel Support** - Works in teams, channels, and DMs
- 🧵 **Thread Support** - Reply in threads or top-level
- 🔒 **DM Policies** - Control who can message the bot
- 👀 **Identity Aware** - Professional enterprise presence
- 💬 **Rich Cards** - Support for adaptive cards (future)
- 📎 **Media Support** - File attachments, images

## Prerequisites

1. **Azure Bot Registration**:
   - Create Azure Bot resource
   - Get App ID and App Password
   - Configure messaging endpoint

2. **Microsoft Teams Setup**:
   - Create Teams app package
   - Install to team/sideload
   - Or publish to org

## Configuration

```yaml
channels:
  msteams:
    appId: "00000000-0000-0000-0000-000000000000"
    appPassword: "secret"
    tenantId: "00000000-0000-0000-0000-000000000000"
    webhookPort: 3978
    webhookPath: "/api/messages"
    requireMention: true
    replyStyle: "thread"
```

## Features

- **Mention-gated** - Only responds to @mentions in channels
- **Thread replies** - Keeps conversations organized
- **Enterprise ready** - Azure AD integration

## Security

- Credentials stored via Azure Bot Framework
- No message content logged (only metadata)
- HTTPS required for webhooks
