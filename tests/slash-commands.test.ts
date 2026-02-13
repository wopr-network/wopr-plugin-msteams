/**
 * Tests for slash command registration and handling (WOP-115)
 *
 * Tests:
 * - Commands registered via channel provider are matched
 * - Slash commands are dispatched with correct args
 * - Non-matching text passes through to normal inject flow
 * - Command handler result is sent back as response
 * - Failed command handler sends error message
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockContext } from "./mocks/wopr-context.js";

const mockSendActivity = vi.fn().mockResolvedValue({});
const mockProcess = vi.fn(async (req: any, res: any, handler: any) => {
  if (req.__activity) {
    await handler({
      activity: req.__activity,
      sendActivity: mockSendActivity,
    });
  }
});

vi.mock("botbuilder", () => {
  return {
    CloudAdapter: class MockCloudAdapter {
      onTurnError: any;
      constructor() { this.onTurnError = null; }
      process = mockProcess;
      continueConversationAsync = vi.fn();
    },
    ConfigurationBotFrameworkAuthentication: class { constructor(config: any) {} },
    TurnContext: class {
      static getConversationReference(activity: any) {
        return { conversation: activity.conversation, bot: activity.recipient };
      }
    },
    CardFactory: {
      adaptiveCard: vi.fn((card: any) => ({
        contentType: "application/vnd.microsoft.card.adaptive",
        content: card,
      })),
    },
    MessageFactory: {
      attachment: vi.fn((a: any) => ({ type: "message", attachments: [a] })),
    },
  };
});

vi.mock("winston", () => {
  const mockLogger = { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() };
  return {
    default: {
      createLogger: vi.fn(() => mockLogger),
      format: { combine: vi.fn(), timestamp: vi.fn(), errors: vi.fn(), json: vi.fn(), colorize: vi.fn(), simple: vi.fn() },
      transports: { File: class { constructor() {} }, Console: class { constructor() {} } },
    },
  };
});

vi.mock("axios", () => ({
  default: { get: vi.fn(), post: vi.fn() },
}));

function makeActivity(overrides: Record<string, any> = {}) {
  return {
    type: "message",
    text: "Hello WOPR",
    from: { id: "user-1", name: "Test User" },
    recipient: { id: "bot-1", name: "WOPR Bot" },
    conversation: { id: "conv-1", conversationType: "personal", name: "Test Chat" },
    ...overrides,
  };
}

describe("slash commands", () => {
  let capturedProvider: any = null;

  beforeEach(() => {
    vi.clearAllMocks();
    capturedProvider = null;
    delete process.env.MSTEAMS_APP_ID;
    delete process.env.MSTEAMS_APP_PASSWORD;
    delete process.env.MSTEAMS_TENANT_ID;
  });

  afterEach(() => {
    vi.resetModules();
  });

  async function initWithCommands(commands: any[] = [], configOverrides: Record<string, any> = {}) {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const mockCtx = createMockContext({
      appId: "test-id",
      appPassword: "test-pass",
      tenantId: "test-tenant",
      dmPolicy: "open",
      useAdaptiveCards: false,
      ...configOverrides,
    });

    // Capture the channel provider when registered
    mockCtx.registerChannelProvider = vi.fn((provider: any) => {
      capturedProvider = provider;
    });

    await plugin.init(mockCtx as any);

    // Register commands via the captured channel provider
    for (const cmd of commands) {
      capturedProvider.registerCommand(cmd);
    }

    return { mod, plugin, mockCtx };
  }

  it("registers commands via channel provider", async () => {
    const handler = vi.fn().mockResolvedValue("done");
    await initWithCommands([{ name: "status", description: "Show status", handler }]);

    const commands = capturedProvider.getCommands();
    expect(commands).toHaveLength(1);
    expect(commands[0].name).toBe("status");
  });

  it("dispatches slash command when text matches /command", async () => {
    const handler = vi.fn().mockResolvedValue("Status: OK");
    const { mod } = await initWithCommands([
      { name: "status", description: "Show status", handler },
    ]);

    const activity = makeActivity({ text: "/status" });
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    expect(handler).toHaveBeenCalledWith(
      expect.objectContaining({
        args: "",
        userId: "user-1",
        userName: "Test User",
      })
    );
  });

  it("passes args to command handler", async () => {
    const handler = vi.fn().mockResolvedValue("Help text");
    const { mod } = await initWithCommands([
      { name: "help", description: "Get help", handler },
    ]);

    const activity = makeActivity({ text: "/help deploy commands" });
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    expect(handler).toHaveBeenCalledWith(
      expect.objectContaining({
        args: "deploy commands",
      })
    );
  });

  it("sends command result back as response", async () => {
    const handler = vi.fn().mockResolvedValue("All systems operational");
    const { mod } = await initWithCommands([
      { name: "status", description: "Show status", handler },
    ]);

    const activity = makeActivity({ text: "/status" });
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    expect(mockSendActivity).toHaveBeenCalledWith(
      expect.objectContaining({
        text: "All systems operational",
      })
    );
  });

  it("passes non-command text to inject flow", async () => {
    const handler = vi.fn().mockResolvedValue("done");
    const { mod, mockCtx } = await initWithCommands([
      { name: "status", description: "Show status", handler },
    ]);

    const activity = makeActivity({ text: "just chatting" });
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    // Should NOT call the command handler
    expect(handler).not.toHaveBeenCalled();
    // Should call inject for normal message flow
    expect(mockCtx.inject).toHaveBeenCalled();
  });

  it("sends error message when command handler throws", async () => {
    const handler = vi.fn().mockRejectedValue(new Error("boom"));
    const { mod } = await initWithCommands([
      { name: "fail", description: "Will fail", handler },
    ]);

    const activity = makeActivity({ text: "/fail" });
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    expect(mockSendActivity).toHaveBeenCalledWith(
      expect.objectContaining({
        text: expect.stringContaining("Command /fail failed"),
      })
    );
  });

  it("unregisters commands correctly", async () => {
    const handler = vi.fn().mockResolvedValue("done");
    await initWithCommands([
      { name: "status", description: "Show status", handler },
    ]);

    expect(capturedProvider.getCommands()).toHaveLength(1);
    capturedProvider.unregisterCommand("status");
    expect(capturedProvider.getCommands()).toHaveLength(0);
  });

  it("does not match partial command names", async () => {
    const handler = vi.fn().mockResolvedValue("done");
    const { mod, mockCtx } = await initWithCommands([
      { name: "status", description: "Show status", handler },
    ]);

    const activity = makeActivity({ text: "/statusbar check" });
    await mod.handleWebhook(
      { __activity: activity },
      { status: vi.fn().mockReturnThis(), send: vi.fn() }
    );

    // /statusbar does not match /status, so it should go to inject
    expect(handler).not.toHaveBeenCalled();
    expect(mockCtx.inject).toHaveBeenCalled();
  });
});
