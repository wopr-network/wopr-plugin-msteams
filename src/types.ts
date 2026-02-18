/**
 * Type definitions for the WOPR MS Teams plugin.
 *
 * These are local definitions compatible with @wopr-network/plugin-types v0.2.0.
 * The installed package may not have dist files, so we define all needed types here.
 */

// ============================================================================
// Core plugin types
// ============================================================================

export interface AgentIdentity {
  name?: string;
  creature?: string;
  vibe?: string;
  emoji?: string;
}

export interface UserProfile {
  name?: string;
  preferredAddress?: string;
  pronouns?: string;
  timezone?: string;
  notes?: string;
}

export interface ChannelRef {
  id: string;
  type: string;
  name?: string;
}

export interface StreamMessage {
  type: "text" | "tool_use" | "complete" | "error" | "system";
  content: string;
  toolName?: string;
  subtype?: string;
  metadata?: Record<string, unknown>;
}

export interface PluginInjectOptions {
  silent?: boolean;
  from?: string;
  channel?: ChannelRef;
  images?: string[];
  source?: unknown;
  contextProviders?: string[];
  priority?: number;
}

export interface PluginLogger {
  info(message: string, ...args: unknown[]): void;
  warn(message: string, ...args: unknown[]): void;
  error(message: string, ...args: unknown[]): void;
  debug(message: string, ...args: unknown[]): void;
}

// ============================================================================
// Config types
// ============================================================================

export interface ConfigField {
  name: string;
  type: "text" | "password" | "number" | "boolean" | "select" | "array" | "object";
  label: string;
  placeholder?: string;
  required?: boolean;
  default?: unknown;
  description?: string;
  options?: Array<{ value: string; label: string }>;
}

export interface ConfigSchema {
  title: string;
  description: string;
  fields: ConfigField[];
}

// ============================================================================
// Channel types
// ============================================================================

export interface ChannelCommandContext {
  channel: string;
  channelType: string;
  sender: string;
  args: string[];
  reply: (msg: string) => Promise<void>;
  getBotUsername: () => string;
}

export interface ChannelMessageContext {
  channel: string;
  channelType: string;
  sender: string;
  content: string;
  reply: (msg: string) => Promise<void>;
  getBotUsername: () => string;
}

export interface ChannelCommand {
  name: string;
  description: string;
  handler: (ctx: ChannelCommandContext) => Promise<void>;
}

export interface ChannelMessageParser {
  id: string;
  pattern: RegExp | ((msg: string) => boolean);
  handler: (ctx: ChannelMessageContext) => Promise<void>;
}

export interface ChannelProvider {
  id: string;
  registerCommand(cmd: ChannelCommand): void;
  unregisterCommand(name: string): void;
  getCommands(): ChannelCommand[];
  addMessageParser(parser: ChannelMessageParser): void;
  removeMessageParser(id: string): void;
  getMessageParsers(): ChannelMessageParser[];
  send(channel: string, content: string): Promise<void>;
  getBotUsername(): string;
}

// ============================================================================
// Plugin manifest
// ============================================================================

export type PluginCategory = "channel" | "llm" | "voice" | "storage" | "tool" | "utility";
export type PluginCapability = "channel" | "llm" | "tts" | "stt" | "image" | "storage";

export interface PluginLifecycle {
  shutdownBehavior?: "graceful" | "immediate";
  shutdownTimeoutMs?: number;
}

export interface PluginManifest {
  name: string;
  version: string;
  description: string;
  author: string;
  license: string;
  capabilities: PluginCapability[];
  category: PluginCategory;
  tags?: string[];
  icon?: string;
  requires?: {
    env?: string[];
    network?: {
      outbound?: boolean;
      inbound?: boolean;
    };
    capabilities?: string[];
  };
  configSchema?: ConfigSchema;
  lifecycle?: PluginLifecycle;
}

// ============================================================================
// Plugin context
// ============================================================================

export interface WOPRPluginContext {
  inject(session: string, message: string, options?: PluginInjectOptions): Promise<string>;
  logMessage(
    session: string,
    message: string,
    options?: { from?: string; senderId?: string; channel?: ChannelRef },
  ): void;
  getAgentIdentity(): AgentIdentity | Promise<AgentIdentity>;
  getUserProfile(): UserProfile | Promise<UserProfile>;
  getSessions(): string[];
  cancelInject(session: string): boolean;
  getConfig<T = unknown>(): T;
  saveConfig<T = unknown>(config: T): Promise<void>;
  getMainConfig(key?: string): unknown;
  registerConfigSchema(pluginId: string, schema: ConfigSchema): void;
  unregisterConfigSchema(pluginId: string): void;
  getConfigSchema(pluginId: string): ConfigSchema | undefined;
  registerExtension(name: string, extension: unknown): void;
  unregisterExtension(name: string): void;
  getExtension<T = unknown>(name: string): T | undefined;
  listExtensions(): string[];
  registerChannelProvider(provider: ChannelProvider): void;
  unregisterChannelProvider(id: string): void;
  getChannelProvider(id: string): ChannelProvider | undefined;
  getChannelProviders(): ChannelProvider[];
  registerProvider(provider: unknown): void;
  unregisterProvider(id: string): void;
  getProvider(id: string): unknown;
  log: PluginLogger;
  getPluginDir(): string;
}

// ============================================================================
// Plugin interface
// ============================================================================

export interface WOPRPlugin {
  name: string;
  version: string;
  description: string;
  manifest?: PluginManifest;
  init(context: WOPRPluginContext): Promise<void>;
  shutdown(): Promise<void>;
}
