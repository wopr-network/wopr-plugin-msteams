/**
 * WOPR Microsoft Teams Plugin - Azure Bot Framework integration
 *
 * Features:
 * - Adaptive Card support for rich responses
 * - Thread/reply support
 * - File attachment handling
 * - Slash command registration
 * - Error retry with exponential backoff + jitter
 */

import path from "node:path";
import winston from "winston";
import axios from "axios";
import {
	CloudAdapter,
	ConfigurationBotFrameworkAuthentication,
	TurnContext,
	Activity,
	ConversationReference,
	CardFactory,
	MessageFactory,
	Attachment,
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
import { createMsteamsExtension, type MsteamsPluginState } from "./msteams-extension";

// ============================================================================
// Types
// ============================================================================

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
	useAdaptiveCards?: boolean;
	maxRetries?: number;
	retryBaseDelayMs?: number;
}

/** Options for building an Adaptive Card. */
export interface AdaptiveCardOptions {
	title?: string;
	body: string;
	actions?: Array<{
		type: "Action.OpenUrl" | "Action.Submit";
		title: string;
		url?: string;
		data?: Record<string, unknown>;
	}>;
	imageUrl?: string;
}

/** Result from downloading a Teams attachment. */
export interface DownloadedAttachment {
	filename: string;
	contentType: string;
	content: Buffer;
}

// ============================================================================
// Module-level state
// ============================================================================

let ctx: WOPRPluginContext | null = null;
let config: MSTeamsConfig = {};
let agentIdentity: AgentIdentity = { name: "WOPR", emoji: "\u{1F440}" };
let adapter: CloudAdapter | null = null;
let isShuttingDown = false;
let logger: winston.Logger;

// Store conversation references for proactive messaging
const conversationReferences: Map<
	string,
	Partial<ConversationReference>
> = new Map();

// Plugin runtime state for WebMCP extension
let pluginState: MsteamsPluginState = {
	initialized: false,
	startedAt: null,
	teams: new Map(),
	channels: new Map(),
	tenants: new Set(),
	messagesProcessed: 0,
	activeConversations: new Set(),
};

// ============================================================================
// Logger
// ============================================================================

function initLogger(): winston.Logger {
	const WOPR_HOME =
		process.env.WOPR_HOME || path.join(process.env.HOME || "~", ".wopr");
	return winston.createLogger({
		level: "debug",
		format: winston.format.combine(
			winston.format.timestamp(),
			winston.format.errors({ stack: true }),
			winston.format.json(),
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
					winston.format.simple(),
				),
				level: "warn",
			}),
		],
	});
}

// ============================================================================
// Retry with Exponential Backoff + Jitter
// ============================================================================

/**
 * Parse the Retry-After header value per RFC 7231.
 * Accepts either a number (delay in seconds) or an HTTP date string.
 * Returns delay in milliseconds, or 0 if unparseable.
 */
export function parseRetryAfter(value: string | undefined | null): number {
	if (!value) return 0;
	const seconds = Number(value);
	if (!Number.isNaN(seconds) && seconds >= 0) {
		return seconds * 1000;
	}
	// Try parsing as HTTP date
	const date = Date.parse(value);
	if (!Number.isNaN(date)) {
		const delayMs = date - Date.now();
		return delayMs > 0 ? delayMs : 0;
	}
	return 0;
}

/**
 * Execute an async function with retry on transient errors (429, 5xx).
 * Uses exponential backoff with full jitter.
 */
export async function withRetry<T>(
	fn: () => Promise<T>,
	opts?: { maxRetries?: number; baseDelayMs?: number },
): Promise<T> {
	const maxRetries = opts?.maxRetries ?? config.maxRetries ?? 3;
	const baseDelayMs = opts?.baseDelayMs ?? config.retryBaseDelayMs ?? 1000;

	let lastError: unknown;
	for (let attempt = 0; attempt <= maxRetries; attempt++) {
		try {
			return await fn();
		} catch (err: any) {
			lastError = err;
			const status = err?.response?.status ?? err?.statusCode ?? 0;
			const isRetryable = status === 429 || (status >= 500 && status < 600);

			if (!isRetryable || attempt === maxRetries) {
				throw err;
			}

			// Exponential backoff with full jitter
			const maxDelay = baseDelayMs * Math.pow(2, attempt);
			const delay = Math.floor(Math.random() * maxDelay);

			// Use Retry-After header if present (seconds or HTTP date per RFC 7231)
			const retryAfter = err?.response?.headers?.["retry-after"];
			const retryAfterMs = parseRetryAfter(retryAfter);
			const actualDelay = Math.max(delay, retryAfterMs);

			logger?.debug(
				`Retry attempt ${attempt + 1}/${maxRetries} after ${actualDelay}ms (status: ${status})`,
			);
			await new Promise((resolve) => setTimeout(resolve, actualDelay));
		}
	}
	throw lastError;
}

// ============================================================================
// Adaptive Card Builder
// ============================================================================

/**
 * Build an Adaptive Card JSON payload from options.
 * Returns a Bot Framework Attachment.
 */
export function buildAdaptiveCard(options: AdaptiveCardOptions): any {
	const body: any[] = [];

	if (options.title) {
		body.push({
			type: "TextBlock",
			text: options.title,
			size: "Large",
			weight: "Bolder",
			wrap: true,
		});
	}

	body.push({
		type: "TextBlock",
		text: options.body,
		wrap: true,
	});

	if (options.imageUrl) {
		body.push({
			type: "Image",
			url: options.imageUrl,
			size: "Auto",
		});
	}

	const card: any = {
		type: "AdaptiveCard",
		$schema: "http://adaptivecards.io/schemas/adaptive-card.json",
		version: "1.4",
		body,
	};

	if (options.actions && options.actions.length > 0) {
		card.actions = options.actions.map((action) => {
			if (action.type === "Action.OpenUrl") {
				return {
					type: "Action.OpenUrl",
					title: action.title,
					url: action.url,
				};
			}
			return {
				type: "Action.Submit",
				title: action.title,
				data: action.data ?? {},
			};
		});
	}

	return CardFactory.adaptiveCard(card);
}

// ============================================================================
// File Attachment Handling
// ============================================================================

/** Maximum attachment download size: 25 MB */
const MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024;

/**
 * Allowed hostname suffixes for attachment downloads.
 * Only Microsoft-controlled domains are permitted to prevent SSRF.
 */
const ALLOWED_DOWNLOAD_HOST_SUFFIXES = [
	".botframework.com",
	".teams.microsoft.com",
	".skype.com",
];

/**
 * Validate that a URL is safe for attachment download.
 * - Must be HTTPS
 * - Hostname must match the Microsoft domain allowlist
 */
export function isAllowedDownloadHost(url: string): boolean {
	let parsed: URL;
	try {
		parsed = new URL(url);
	} catch {
		return false;
	}

	if (parsed.protocol !== "https:") {
		return false;
	}

	const hostname = parsed.hostname.toLowerCase();
	return ALLOWED_DOWNLOAD_HOST_SUFFIXES.some(
		(suffix) => hostname === suffix.slice(1) || hostname.endsWith(suffix),
	);
}

/**
 * Check if a URL belongs to a Skype domain using proper URL parsing.
 * Used to determine whether to attach a bot auth token.
 */
function isSkypeDomain(url: string): boolean {
	try {
		const hostname = new URL(url).hostname.toLowerCase();
		return hostname === "skype.com" || hostname.endsWith(".skype.com");
	} catch {
		return false;
	}
}

/**
 * Download an attachment from a Teams message.
 * Uses the Bot Connector API with auth token.
 *
 * Security: Only downloads from allowed Microsoft domains (HTTPS).
 * Enforces a 25 MB size limit to prevent denial-of-service.
 */
export async function downloadAttachment(
	activity: Activity,
	attachmentIndex: number = 0,
): Promise<DownloadedAttachment | null> {
	const attachments = activity.attachments;
	if (!attachments || attachmentIndex >= attachments.length) {
		return null;
	}

	const att = attachments[attachmentIndex];
	const downloadUrl = att.contentUrl;
	if (!downloadUrl) {
		logger?.warn("Attachment has no contentUrl");
		return null;
	}

	if (!isAllowedDownloadHost(downloadUrl)) {
		logger?.warn(
			`Blocked attachment download from disallowed host: ${downloadUrl}`,
		);
		return null;
	}

	return withRetry(async () => {
		const response = await axios.get(downloadUrl, {
			responseType: "arraybuffer",
			maxContentLength: MAX_ATTACHMENT_SIZE,
			headers: isSkypeDomain(downloadUrl)
				? { Authorization: `Bearer ${await getBotToken()}` }
				: {},
		});

		return {
			filename: att.name || "attachment",
			contentType: att.contentType || "application/octet-stream",
			content: Buffer.from(response.data),
		};
	});
}

/**
 * Validate that a URL uses the HTTPS protocol.
 * Prevents mixed content and SSRF via non-HTTPS schemes.
 */
export function isHttpsUrl(url: string): boolean {
	try {
		return new URL(url).protocol === "https:";
	} catch {
		return false;
	}
}

/**
 * Build a file info card for sending a file to the user.
 * In Teams, files are sent via consent/info cards.
 */
export function buildFileCard(
	filename: string,
	contentUrl: string,
	fileSize?: number,
): any {
	if (contentUrl && !isHttpsUrl(contentUrl)) {
		throw new Error(
			`contentUrl must be an HTTPS URL, got: ${contentUrl}`,
		);
	}

	const card: Record<string, any> = {
		contentType: "application/vnd.microsoft.teams.card.file.info",
		name: filename,
		content: {
			fileType: filename.split(".").pop() || "unknown",
			uniqueId: `file-${Date.now()}`,
		},
	};
	if (contentUrl) {
		card.contentUrl = contentUrl;
	}
	if (fileSize !== undefined) {
		(card.content as any).fileSize = fileSize;
	}
	return card;
}

/** Get a bot token for authenticated API calls. */
async function getBotToken(): Promise<string> {
	const creds = resolveCredentials();
	if (!creds) {
		throw new Error(
			"Cannot obtain bot token: MS Teams credentials are not configured. " +
				"Set appId, appPassword, and tenantId in config or environment variables.",
		);
	}

	const tokenResponse = await axios.post(
		`https://login.microsoftonline.com/${creds.tenantId}/oauth2/v2.0/token`,
		new URLSearchParams({
			grant_type: "client_credentials",
			client_id: creds.appId,
			client_secret: creds.appPassword,
			scope: "https://api.botframework.com/.default",
		}).toString(),
		{ headers: { "Content-Type": "application/x-www-form-urlencoded" } },
	);

	const token = tokenResponse.data.access_token;
	if (!token) {
		throw new Error(
			"Bot token response did not contain an access_token.",
		);
	}
	return token;
}

// ============================================================================
// Config Schema
// ============================================================================

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
		{
			name: "useAdaptiveCards",
			type: "boolean",
			label: "Use Adaptive Cards",
			default: true,
			description: "Send rich responses using Adaptive Cards",
		},
		{
			name: "maxRetries",
			type: "number",
			label: "Max Retries",
			placeholder: "3",
			default: 3,
			description: "Maximum retry attempts for failed API calls",
		},
	],
};

// ============================================================================
// Plugin Manifest
// ============================================================================

const manifest: PluginManifest = {
	name: "@wopr-network/wopr-plugin-msteams",
	version: "1.0.0",
	description: "Microsoft Teams integration using Azure Bot Framework",
	author: "WOPR Network",
	license: "MIT",
	capabilities: ["channel"],
	category: "channel",
	tags: [
		"msteams",
		"teams",
		"azure",
		"bot-framework",
		"chat",
		"adaptive-cards",
	],
	icon: "\u{1F7E6}",
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
// Channel Provider
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

	async send(channel: string, content: string): Promise<void> {
		const ref = conversationReferences.get(channel);
		if (!ref || !adapter) {
			logger?.info(
				`No conversation reference for channel ${channel}, cannot send proactively`,
			);
			return;
		}

		await withRetry(async () => {
			await adapter!.continueConversationAsync(
				resolveCredentials()?.appId || "",
				ref as ConversationReference,
				async (turnContext: TurnContext) => {
					if (config.useAdaptiveCards !== false) {
						const card = buildAdaptiveCard({ body: content });
						await turnContext.sendActivity({ attachments: [card] });
					} else {
						await turnContext.sendActivity(content);
					}
				},
			);
		});
	},

	getBotUsername(): string {
		return agentIdentity.name || "WOPR";
	},
};

// ============================================================================
// Slash Command Processing
// ============================================================================

/**
 * Check if a message text matches a registered slash command.
 * Returns the command and parsed args if matched.
 */
function matchSlashCommand(
	text: string,
): { command: ChannelCommand; args: string } | null {
	const trimmed = text.trim();
	for (const [name, cmd] of registeredCommands) {
		const prefix = `/${name}`;
		if (trimmed === prefix || trimmed.startsWith(`${prefix} `)) {
			const args = trimmed.slice(prefix.length).trim();
			return { command: cmd, args };
		}
	}
	return null;
}

// ============================================================================
// Helper Functions
// ============================================================================

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

function resolveCredentials(): {
	appId: string;
	appPassword: string;
	tenantId: string;
} | null {
	const appId = config.appId || process.env.MSTEAMS_APP_ID;
	const appPassword = config.appPassword || process.env.MSTEAMS_APP_PASSWORD;
	const tenantId = config.tenantId || process.env.MSTEAMS_TENANT_ID;

	if (!appId || !appPassword || !tenantId) {
		return null;
	}

	return { appId, appPassword, tenantId };
}

function isAllowed(userId: string, conversationType: string): boolean {
	const isGroup =
		conversationType === "channel" || conversationType === "groupChat";

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

/** Store a conversation reference for proactive messaging. */
function storeConversationReference(activity: Activity): void {
	const key = activity.conversation?.id;
	if (!key) return;

	try {
		const ref = TurnContext.getConversationReference(activity);
		if (ref) {
			conversationReferences.set(key, ref);
		}
	} catch {
		// Fallback: build a minimal reference from the activity
		conversationReferences.set(key, {
			channelId: activity.channelId,
			serviceUrl: activity.serviceUrl,
			conversation: activity.conversation,
			bot: activity.recipient,
		} as Partial<ConversationReference>);
	}
}

// ============================================================================
// Activity Processing
// ============================================================================

async function processActivity(turnContext: TurnContext): Promise<void> {
	if (!ctx) return;

	const activity = turnContext.activity;

	// Store conversation reference for proactive messaging
	storeConversationReference(activity);

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
				(e) =>
					e.type === "mention" && e.mentioned?.id === activity.recipient?.id,
			);
			if (!mentioned) {
				logger.debug("Skipping message without mention in group");
				return;
			}
		}
	}

	// Track state for WebMCP extension
	pluginState.messagesProcessed++;
	pluginState.activeConversations.add(conversationId);

	// Track tenant
	const tenantId = activity.conversation?.tenantId;
	if (tenantId) {
		pluginState.tenants.add(tenantId);
	}

	// Track team info from channelData
	const teamData = (activity as any).channelData?.team;
	if (teamData?.id) {
		pluginState.teams.set(teamData.id, {
			id: teamData.id,
			name: teamData.name || teamData.id,
		});
	}

	// Track channel info
	const channelData = (activity as any).channelData?.channel;
	if (channelData?.id && teamData?.id) {
		if (!pluginState.channels.has(teamData.id)) {
			pluginState.channels.set(teamData.id, new Map());
		}
		pluginState.channels.get(teamData.id)!.set(channelData.id, {
			id: channelData.id,
			name: channelData.name || activity.conversation?.name || channelData.id,
			type: channelData.type || "standard",
		});
	}

	// Check for slash commands first
	const cmdMatch = matchSlashCommand(text);
	if (cmdMatch) {
		await handleSlashCommand(
			turnContext,
			cmdMatch.command,
			cmdMatch.args,
			userId,
			userName,
			conversationId,
		);
		return;
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

	// Handle file attachments if present
	if (activity.attachments && activity.attachments.length > 0) {
		const fileAttachments = activity.attachments.filter(
			(a) =>
				a.contentType !== "text/html" &&
				!a.contentType?.startsWith("application/vnd.microsoft.card"),
		);
		if (fileAttachments.length > 0) {
			const fileNames = fileAttachments.map((a) => a.name || "file").join(", ");
			logger.info(
				`Received ${fileAttachments.length} file attachment(s): ${fileNames}`,
			);
		}
	}

	// Inject to WOPR
	await injectMessage(turnContext, text, userName, sessionKey, channelInfo);
}

async function handleSlashCommand(
	turnContext: TurnContext,
	command: ChannelCommand,
	args: string,
	userId: string,
	_userName: string,
	conversationId: string,
): Promise<void> {
	try {
		await command.handler({
			channel: `msteams:${conversationId}`,
			channelType: "msteams",
			sender: userId,
			args: args ? args.split(" ").filter(Boolean) : [],
			reply: async (msg: string) => { await sendResponse(turnContext, msg); },
			getBotUsername: () => agentIdentity.name || "WOPR",
		});
	} catch (err) {
		logger.error(`Slash command /${command.name} failed:`, err);
		await sendResponse(
			turnContext,
			`Command /${command.name} failed. Please try again.`,
		);
	}
}

async function injectMessage(
	turnContext: TurnContext,
	text: string,
	userName: string,
	sessionKey: string,
	channelInfo: ChannelRef,
): Promise<void> {
	if (!ctx) return;

	const prefix = `[${userName}]: `;
	const messageWithPrefix = prefix + text;

	const response = await ctx.inject(sessionKey, messageWithPrefix, {
		from: userName,
		channel: channelInfo,
	});

	// Send response back
	await sendResponse(turnContext, response);
}

// ============================================================================
// Response Sending
// ============================================================================

async function sendResponse(
	turnContext: TurnContext,
	text: string,
): Promise<void> {
	const activity = turnContext.activity;

	await withRetry(async () => {
		if (config.useAdaptiveCards !== false) {
			// Send as adaptive card for rich formatting
			const card = buildAdaptiveCard({ body: text });
			const replyActivity = MessageFactory.attachment(card);

			// Thread reply support
			if (config.replyStyle === "thread" && activity.id) {
				replyActivity.replyToId = activity.id;
			}

			await turnContext.sendActivity(replyActivity);
		} else {
			// Plain text/markdown response
			const reply: Partial<Activity> = {
				type: "message",
				text,
				textFormat: "markdown",
			};

			if (config.replyStyle === "thread" && activity.id) {
				reply.replyToId = activity.id;
			}

			await turnContext.sendActivity(reply);
		}
	});
}

// ============================================================================
// Webhook Handler
// ============================================================================

export async function handleWebhook(req: any, res: any): Promise<void> {
	if (!adapter) {
		res.status(500).send("Bot not initialized");
		return;
	}

	await adapter.process(req, res, async (context) => {
		await processActivity(context);
	});
}

// ============================================================================
// Adapter Initialization
// ============================================================================

function initAdapter(): CloudAdapter | null {
	const creds = resolveCredentials();
	if (!creds) return null;

	const botFrameworkAuthentication =
		new ConfigurationBotFrameworkAuthentication({
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

// ============================================================================
// Extension API
// ============================================================================

const msteamsExtension = {
	getBotUsername: () => agentIdentity.name || "WOPR",
	handleWebhook,
	/** Send an adaptive card to a conversation. */
	sendAdaptiveCard: async (
		conversationId: string,
		options: AdaptiveCardOptions,
	): Promise<void> => {
		const ref = conversationReferences.get(conversationId);
		if (!ref || !adapter) {
			logger?.warn(
				`Cannot send adaptive card: no reference for ${conversationId}`,
			);
			return;
		}

		await withRetry(async () => {
			await adapter!.continueConversationAsync(
				resolveCredentials()?.appId || "",
				ref as ConversationReference,
				async (turnContext: TurnContext) => {
					const card = buildAdaptiveCard(options);
					await turnContext.sendActivity({ attachments: [card] });
				},
			);
		});
	},
	/** Download an attachment from a message. */
	downloadAttachment,
	/** Get stored conversation references. */
	getConversationReferences: () => new Map(conversationReferences),
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
				"MS Teams credentials not configured. Run 'wopr configure --plugin msteams' to set up.",
			);
			return;
		}

		// Validate that credentials are non-empty strings
		if (!creds.appId.trim()) {
			throw new Error(
				"MS Teams appId is empty. Provide a valid Azure Bot App ID.",
			);
		}
		if (!creds.appPassword.trim()) {
			throw new Error(
				"MS Teams appPassword is empty. Provide a valid Azure Bot App Password.",
			);
		}
		if (!creds.tenantId.trim()) {
			throw new Error(
				"MS Teams tenantId is empty. Provide a valid Azure AD Tenant ID.",
			);
		}

		// Initialize adapter
		adapter = initAdapter();
		if (!adapter) {
			logger.error("Failed to initialize MS Teams adapter");
			return;
		}

		// Mark state as initialized
		pluginState.initialized = true;
		pluginState.startedAt = Date.now();

		// Create and register the WebMCP extension
		const webmcpExtension = createMsteamsExtension(
			() => adapter,
			() => ctx,
			() => pluginState,
		);

		if (ctx.registerExtension) {
			ctx.registerExtension("msteams-webmcp", webmcpExtension);
			logger.info("Registered MS Teams WebMCP extension");
		}

		logger.info("MS Teams plugin initialized");
		logger.info(
			`Webhook endpoint: http://localhost:${config.webhookPort || 3978}${config.webhookPath || "/api/messages"}`,
		);
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
		conversationReferences.clear();
		adapter = null;
		ctx = null;

		// Reset plugin state
		pluginState = {
			initialized: false,
			startedAt: null,
			teams: new Map(),
			channels: new Map(),
			tenants: new Set(),
			messagesProcessed: 0,
			activeConversations: new Set(),
		};
	},
};

export default plugin;
