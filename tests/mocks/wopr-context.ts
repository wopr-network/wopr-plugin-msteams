/**
 * Mock WOPRPluginContext for testing wopr-plugin-msteams.
 */
import { vi } from "vitest";

export interface MockWOPRPluginContext {
  inject: ReturnType<typeof vi.fn>;
  logMessage: ReturnType<typeof vi.fn>;
  getAgentIdentity: ReturnType<typeof vi.fn>;
  getConfig: ReturnType<typeof vi.fn>;
  saveConfig: ReturnType<typeof vi.fn>;
  registerConfigSchema: ReturnType<typeof vi.fn>;
  getPluginDir: ReturnType<typeof vi.fn>;
  log: {
    info: ReturnType<typeof vi.fn>;
    warn: ReturnType<typeof vi.fn>;
    error: ReturnType<typeof vi.fn>;
  };
}

export function createMockContext(
  configData: Record<string, any> = {}
): MockWOPRPluginContext {
  return {
    inject: vi.fn().mockResolvedValue("Mock response"),
    logMessage: vi.fn(),
    getAgentIdentity: vi.fn().mockResolvedValue({
      name: "TestAgent",
      creature: "owl",
      vibe: "chill",
      emoji: "🦉",
    }),
    getConfig: vi.fn(() => ({ ...configData })),
    saveConfig: vi.fn(async (config: any) => {
      Object.assign(configData, config);
    }),
    registerConfigSchema: vi.fn(),
    getPluginDir: vi.fn().mockReturnValue("/tmp/wopr-test/plugins/msteams"),
    log: {
      info: vi.fn(),
      warn: vi.fn(),
      error: vi.fn(),
    },
  };
}
