/**
 * MS Teams Extension (cross-plugin API)
 *
 * Provides data-access methods for WebMCP tools to read MS Teams
 * bot connection state, teams, channels, and message stats.
 *
 * Bot Framework is webhook-based (not WebSocket), so state is
 * accumulated passively from incoming activities.
 */

// ============================================================================
// Runtime state interface
// ============================================================================

/**
 * Runtime state tracked by the plugin from incoming Bot Framework activities.
 * Since CloudAdapter is webhook-based, there is no persistent connection to
 * poll — all state is accumulated passively as messages arrive.
 */
export interface MsteamsPluginState {
  /** Whether the adapter is initialized and ready */
  initialized: boolean;
  /** Timestamp when adapter was initialized */
  startedAt: number | null;
  /** Teams the bot has interacted with (populated from incoming activities) */
  teams: Map<string, { id: string; name: string }>;
  /** Channels the bot has interacted with, keyed by teamId */
  channels: Map<string, Map<string, { id: string; name: string; type: string }>>;
  /** Tenant IDs the bot has received messages from */
  tenants: Set<string>;
  /** Total messages processed */
  messagesProcessed: number;
  /** Total number of unique conversations ever seen */
  totalConversations: number;
}

// ============================================================================
// Structured return types for WebMCP-facing extension methods
// ============================================================================

export interface MsteamsStatusInfo {
  online: boolean;
  connectedTenants: number;
  /** Always -1 — Bot Framework is webhook-based, no persistent ping available */
  latencyMs: number;
  uptimeMs: number | null;
}

export interface TeamInfo {
  id: string;
  name: string;
}

export interface MsteamsChannelInfo {
  id: string;
  name: string;
  type: string;
}

export interface MsteamsMessageStatsInfo {
  messagesProcessed: number;
  activeConversations: number;
}

// ============================================================================
// Extension interface
// ============================================================================

export interface MsteamsExtension {
  getStatus: () => MsteamsStatusInfo;
  listTeams: () => TeamInfo[];
  listChannels: (teamId?: string) => MsteamsChannelInfo[];
  getMessageStats: () => MsteamsMessageStatsInfo;
}

// ============================================================================
// Factory function
// ============================================================================

export function createMsteamsExtension(
  getState: () => MsteamsPluginState,
): MsteamsExtension {
  return {
    getStatus: (): MsteamsStatusInfo => {
      const state = getState();
      return {
        online: state.initialized,
        connectedTenants: state.tenants.size,
        latencyMs: -1,
        uptimeMs: state.startedAt !== null ? Date.now() - state.startedAt : null,
      };
    },

    listTeams: (): TeamInfo[] => {
      const state = getState();
      return Array.from(state.teams.values()).map((t) => ({
        id: t.id,
        name: t.name,
      }));
    },

    listChannels: (teamId?: string): MsteamsChannelInfo[] => {
      const state = getState();
      if (teamId) {
        const teamChannels = state.channels.get(teamId);
        if (!teamChannels) return [];
        return Array.from(teamChannels.values()).map((ch) => ({
          id: ch.id,
          name: ch.name,
          type: ch.type,
        }));
      }
      // Return all channels across all teams
      const all: MsteamsChannelInfo[] = [];
      for (const teamChannels of state.channels.values()) {
        for (const ch of teamChannels.values()) {
          all.push({ id: ch.id, name: ch.name, type: ch.type });
        }
      }
      return all;
    },

    getMessageStats: (): MsteamsMessageStatsInfo => {
      const state = getState();
      return {
        messagesProcessed: state.messagesProcessed,
        activeConversations: state.totalConversations,
      };
    },
  };
}
