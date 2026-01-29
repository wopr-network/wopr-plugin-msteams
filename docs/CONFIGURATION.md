# MS Teams Configuration

## Required Settings

| Option | Type | Required | Description |
|--------|------|----------|-------------|
| `appId` | string | **Yes** | Azure Bot App ID |
| `appPassword` | string | **Yes** | Azure Bot App Password |
| `tenantId` | string | **Yes** | Azure AD Tenant ID |
| `messagingEndpoint` | string | **Yes** | Public HTTPS URL + /api/messages |

## Optional Settings

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `channelPolicy` | string | `"mention"` | Respond to: "mention", "all", "none" |
| `dmPolicy` | string | `"open"` | DM handling mode |
| `useTeamsChannel` | boolean | `true` | Create thread per conversation |

## Setup Steps

1. **Create Azure Bot**
   - Go to Azure Portal
   - Create "Azure Bot" resource
   - Note App ID and create App Password

2. **Configure Messaging Endpoint**
   - Set endpoint to `https://your-server.com/api/messages`
   - Must be HTTPS and publicly accessible

3. **Create Teams App**
   - Go to Teams Developer Portal
   - Create app, add Bot capability
   - Enter Azure Bot App ID
   - Download and install app package

## Configuration Example

```json
{
  "appId": "your-app-id",
  "appPassword": "your-app-password",
  "tenantId": "your-tenant-id",
  "messagingEndpoint": "https://wopr.example.com/api/messages",
  "channelPolicy": "mention",
  "dmPolicy": "open"
}
```

## Environment Variables

```bash
export TEAMS_APP_ID="..."
export TEAMS_APP_PASSWORD="..."
export TEAMS_TENANT_ID="..."
```
