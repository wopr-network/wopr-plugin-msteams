/**
 * WOPR Microsoft Teams Plugin - Azure Bot Framework integration
 */

import path from "node:path";
import winston from "winston";
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
  Activity,
} from "botbuilder";
import type {
  WOPRPlugin,
  WOPRPluginContext,
  ConfigSchema,
  AgentIdentity,
  ChannelRef,
  ChannelCommand,
  ChannelMessageParser,
  ChannelProvider,
  PluginManifest,
} from "./types.js";

// MS Teams config interface
interface MSTeamsConfig {
  appId?: string;
  appPassword?: string;
  tenantId?: string;
  enabled?: boolean;
  webhookPort?: number;
  webhookPath?: string;
  dmPolicy?: "allowlist" | "pairing" | "open" | "disabled";
  allowFrom?: string[];
  groupPolicy?: "allowlist" | "open" | "disabled";
  groupAllowFrom?: string[];
  requireMention?: boolean;
  replyStyle?: "thread" | "top-level";
}

// Module-level state
let ctx: WOPRPluginContext | null = null;
let config: MSTeamsConfig = {};
let agentIdentity: AgentIdentity = { name: "WOPR", emoji: "👀" };
let adapter: CloudAdapter | null = null;
let isShuttingDown = false;
let logger: winston.Logger;

// Initialize winston logger
function initLogger(): winston.Logger {
  const WOPR_HOME = process.env.WOPR_HOME || path.join(process.env.HOME || "~", ".wopr");
  return winston.createLogger({
    level: "debug",
    format: winston.format.combine(
      winston.format.timestamp(),
      winston.format.errors({ stack: true }),
      winston.format.json()
    ),
    defaultMeta: { service: "wopr-plugin-msteams" },
    transports: [
      new winston.transports.File({
        filename: path.join(WOPR_HOME, "logs", "msteams-plugin-error.log"),
        level: "error",
      }),
      new winston.transports.File({
        filename: path.join(WOPR_HOME, "logs", "msteams-plugin.log"),
        level: "debug",
      }),
      new winston.transports.Console({
        format: winston.format.combine(
          winston.format.colorize(),
          winston.format.simple()
        ),
        level: "warn",
      }),
    ],
  });
}

// Config schema
const configSchema: ConfigSchema = {
  title: "Microsoft Teams Integration",
  description: "Configure Microsoft Teams Bot using Azure Bot Framework",
  fields: [
    {
      name: "appId",
      type: "text",
      label: "App ID",
      placeholder: "00000000-0000-0000-0000-000000000000",
      required: true,
      description: "Azure Bot App ID",
    },
    {
      name: "appPassword",
      type: "password",
      label: "App Password",
      placeholder: "secret",
      required: true,
      description: "Azure Bot App Password (Client Secret)",
    },
    {
      name: "tenantId",
      type: "text",
      label: "Tenant ID",
      placeholder: "00000000-0000-0000-0000-000000000000",
      required: true,
      description: "Azure AD Tenant ID",
    },
    {
      name: "webhookPort",
      type: "number",
      label: "Webhook Port",
      placeholder: "3978",
      default: 3978,
      description: "Port for webhook server",
    },
    {
      name: "webhookPath",
      type: "text",
      label: "Webhook Path",
      placeholder: "/api/messages",
      default: "/api/messages",
      description: "Path for webhook endpoint",
    },
    {
      name: "dmPolicy",
      type: "select",
      label: "DM Policy",
      placeholder: "pairing",
      default: "pairing",
      description: "How to handle direct messages",
    },
    {
      name: "allowFrom",
      type: "array",
      label: "Allowed Users",
      placeholder: "user-id-1, user-id-2",
      description: "Allowed user IDs for DMs",
    },
    {
      name: "groupPolicy",
      type: "select",
      label: "Group Policy",
      placeholder: "allowlist",
      default: "allowlist",
      description: "How to handle channel/group messages",
    },
    {
      name: "requireMention",
      type: "boolean",
      label: "Require Mention",
      default: true,
      description: "Require @mention in channels",
    },
    {
      name: "replyStyle",
      type: "select",
      label: "Reply Style",
      placeholder: "thread",
      default: "thread",
      description: "Reply in thread or top-level",
    },
  ],
};

// ============================================================================
// Plugin Manifest (WaaS metadata)
// ============================================================================

const manifest: PluginManifest = {
  name: "@wopr-network/wopr-plugin-msteams",
  version: "1.0.0",
  description: "Microsoft Teams integration using Azure Bot Framework",
  author: "WOPR Network",
  license: "MIT",
  capabilities: ["channel"],
  category: "channel",
  tags: ["msteams", "teams", "azure", "bot-framework", "chat"],
  icon: "🟦",
  requires: {
    env: ["MSTEAMS_APP_ID", "MSTEAMS_APP_PASSWORD", "MSTEAMS_TENANT_ID"],
    network: {
      outbound: true,
      inbound: true,
    },
  },
  configSchema,
  lifecycle: {
    shutdownBehavior: "graceful",
    shutdownTimeoutMs: 10000,
  },
};

// ============================================================================
// Channel Provider (cross-plugin command/parser registration)
// ============================================================================

const registeredCommands: Map<string, ChannelCommand> = new Map();
const registeredParsers: Map<string, ChannelMessageParser> = new Map();

const msteamsChannelProvider: ChannelProvider = {
  id: "msteams",

  registerCommand(cmd: ChannelCommand): void {
    registeredCommands.set(cmd.name, cmd);
    logger?.info(`Channel command registered: ${cmd.name}`);
  },

  unregisterCommand(name: string): void {
    registeredCommands.delete(name);
  },

  getCommands(): ChannelCommand[] {
    return Array.from(registeredCommands.values());
  },

  addMessageParser(parser: ChannelMessageParser): void {
    registeredParsers.set(parser.id, parser);
    logger?.info(`Message parser registered: ${parser.id}`);
  },

  removeMessageParser(id: string): void {
    registeredParsers.delete(id);
  },

  getMessageParsers(): ChannelMessageParser[] {
    return Array.from(registeredParsers.values());
  },

  async send(_channel: string, content: string): Promise<void> {
    // MS Teams requires a TurnContext to send proactive messages.
    // Proactive messaging requires storing conversation references
    // which is handled by the webhook flow. For now, log the intent.
    logger?.info(`Channel send requested: ${content.substring(0, 100)}...`);
  },

  getBotUsername(): string {
    return agentIdentity.name || "WOPR";
  },
};

// ============================================================================
// Helper Functions
// ============================================================================

// Refresh identity
async function refreshIdentity(): Promise<void> {
  if (!ctx) return;
  try {
    const identity = await ctx.getAgentIdentity();
    if (identity) {
      agentIdentity = { ...agentIdentity, ...identity };
      logger.info("Identity refreshed:", agentIdentity.name);
    }
  } catch (e) {
    logger.warn("Failed to refresh identity:", String(e));
  }
}

// Resolve credentials
function resolveCredentials(): { appId: string; appPassword: string; tenantId: string } | null {
  const appId = config.appId || process.env.MSTEAMS_APP_ID;
  const appPassword = config.appPassword || process.env.MSTEAMS_APP_PASSWORD;
  const tenantId = config.tenantId || process.env.MSTEAMS_TENANT_ID;

  if (!appId || !appPassword || !tenantId) {
    return null;
  }

  return { appId, appPassword, tenantId };
}

// Check if sender is allowed
function isAllowed(userId: string, conversationType: string): boolean {
  const isGroup = conversationType === "channel" || conversationType === "groupChat";

  if (isGroup) {
    const policy = config.groupPolicy || "allowlist";
    if (policy === "open") return true;
    if (policy === "disabled") return false;

    const allowed = config.groupAllowFrom || config.allowFrom || [];
    if (allowed.includes("*")) return true;
    return allowed.includes(userId);
  } else {
    const policy = config.dmPolicy || "pairing";
    if (policy === "open") return true;
    if (policy === "disabled") return false;
    if (policy === "pairing") return true;

    const allowed = config.allowFrom || [];
    if (allowed.includes("*")) return true;
    return allowed.includes(userId);
  }
}

// Process incoming activity
async function processActivity(activity: Activity): Promise<void> {
  if (!ctx) return;

  // Skip non-message activities
  if (activity.type !== "message") return;

  // Skip messages from the bot itself
  if (activity.from?.id === activity.recipient?.id) return;

  const userId = activity.from?.id;
  const userName = activity.from?.name || "Unknown";
  const text = activity.text || "";
  const conversationId = activity.conversation?.id;
  const conversationType = activity.conversation?.conversationType;

  if (!userId || !conversationId) return;

  // Check if allowed
  if (!isAllowed(userId, conversationType || "personal")) {
    logger.info(`Message from ${userId} blocked by policy`);
    return;
  }

  // Check for mention requirement in groups
  if (conversationType === "channel" || conversationType === "groupChat") {
    if (config.requireMention !== false) {
      const mentioned = activity.entities?.some(
        (e) => e.type === "mention" && e.mentioned?.id === activity.recipient?.id
      );
      if (!mentioned) {
        logger.debug("Skipping message without mention in group");
        return;
      }
    }
  }

  // Build channel info
  const channelId = `msteams:${conversationId}`;
  const channelInfo: ChannelRef = {
    type: "msteams",
    id: channelId,
    name: activity.conversation?.name || "MS Teams",
  };

  // Log for context
  const sessionKey = `msteams-${conversationId}`;
  ctx.logMessage(sessionKey, text, {
    from: userName,
    channel: channelInfo,
  });

  // Inject to WOPR
  await injectMessage(text, userName, sessionKey, channelInfo, activity);
}

// Inject message to WOPR
async function injectMessage(
  text: string,
  userName: string,
  sessionKey: string,
  channelInfo: ChannelRef,
  activity: Activity
): Promise<void> {
  if (!ctx) return;

  const prefix = `[${userName}]: `;
  const messageWithPrefix = prefix + text;

  const response = await ctx.inject(sessionKey, messageWithPrefix, {
    from: userName,
    channel: channelInfo,
  });

  // Send response back
  await sendResponse(activity, response);
}

// Send response back to MS Teams
async function sendResponse(activity: Activity, text: string): Promise<void> {
  if (!adapter) return;

  // Create reply activity
  const reply: Partial<Activity> = {
    type: "message",
    text,
    textFormat: "markdown",
  };

  // Reply in thread if configured
  if (config.replyStyle === "thread" && activity.id) {
    reply.replyToId = activity.id;
  }

  try {
    // We need a TurnContext to send the response
    // This is typically handled by the webhook handler
    // For now, we'll store the response to be sent by the handler
    logger.info(`Would send response: ${text.substring(0, 100)}...`);
  } catch (err) {
    logger.error("Failed to send MS Teams response:", err);
  }
}

// Webhook handler
export async function handleWebhook(req: any, res: any): Promise<void> {
  if (!adapter) {
    res.status(500).send("Bot not initialized");
    return;
  }

  await adapter.process(req, res, async (context) => {
    await processActivity(context.activity);
  });
}

// Initialize bot adapter
function initAdapter(): CloudAdapter | null {
  const creds = resolveCredentials();
  if (!creds) return null;

  const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
    MicrosoftAppId: creds.appId,
    MicrosoftAppPassword: creds.appPassword,
    MicrosoftAppTenantId: creds.tenantId,
  });

  const newAdapter = new CloudAdapter(botFrameworkAuthentication);

  newAdapter.onTurnError = async (context, error) => {
    logger.error("MS Teams turn error:", error);
    await context.sendActivity("Sorry, something went wrong!");
  };

  return newAdapter;
}

// Extension API exposed to other plugins
const msteamsExtension = {
  getBotUsername: () => agentIdentity.name || "WOPR",
  handleWebhook,
};

// ============================================================================
// Plugin Definition
// ============================================================================

const plugin: WOPRPlugin = {
  name: "msteams",
  version: "1.0.0",
  description: "Microsoft Teams integration using Azure Bot Framework",
  manifest,

  async init(context: WOPRPluginContext): Promise<void> {
    ctx = context;
    config = (context.getConfig() || {}) as MSTeamsConfig;

    // Initialize logger
    logger = initLogger();

    // Register config schema
    ctx.registerConfigSchema("msteams", configSchema);

    // Register as a channel provider so other plugins can add commands/parsers
    if (ctx.registerChannelProvider) {
      ctx.registerChannelProvider(msteamsChannelProvider);
      logger.info("Registered MS Teams channel provider");
    }

    // Register the MS Teams extension so other plugins can interact
    if (ctx.registerExtension) {
      ctx.registerExtension("msteams", msteamsExtension);
      logger.info("Registered MS Teams extension");
    }

    // Refresh identity
    await refreshIdentity();

    // Validate config
    const creds = resolveCredentials();
    if (!creds) {
      logger.warn(
        "MS Teams credentials not configured. Run 'wopr configure --plugin msteams' to set up."
      );
      return;
    }

    // Initialize adapter
    adapter = initAdapter();
    if (!adapter) {
      logger.error("Failed to initialize MS Teams adapter");
      return;
    }

    logger.info("MS Teams plugin initialized");
    logger.info(`Webhook endpoint: http://localhost:${config.webhookPort || 3978}${config.webhookPath || "/api/messages"}`);
    logger.info("Make sure to register this URL in Azure Bot Configuration");
  },

  async shutdown(): Promise<void> {
    isShuttingDown = true;

    logger.info("Shutting down MS Teams plugin...");

    // Unregister channel provider and extension
    if (ctx?.unregisterChannelProvider) {
      ctx.unregisterChannelProvider("msteams");
    }
    if (ctx?.unregisterExtension) {
      ctx.unregisterExtension("msteams");
    }

    registeredCommands.clear();
    registeredParsers.clear();
    adapter = null;
    ctx = null;
  },
};

export default plugin;
