/**
 * Type re-exports from the shared @wopr-network/plugin-types package,
 * plus plugin-specific types for MS Teams.
 */

// Re-export shared types used by this plugin
export type {
  AgentIdentity,
  ChannelCommand,
  ChannelCommandContext,
  ChannelMessageContext,
  ChannelMessageParser,
  ChannelProvider,
  ChannelRef,
  ConfigSchema,
  PluginCapability,
  PluginCategory,
  PluginLifecycle,
  PluginLogger,
  PluginManifest,
  StreamMessage,
  UserProfile,
  WOPRPlugin,
  WOPRPluginContext,
} from "@wopr-network/plugin-types";
