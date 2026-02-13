/**
 * Tests for MS Teams message handling logic (WOP-243, WOP-115)
 *
 * Tests:
 * - handleWebhook routes activities through the adapter
 * - Message processing invokes ctx.inject for valid messages
 * - Bot's own messages are skipped
 * - Non-message activities are skipped
 * - DM policy enforcement (allowlist, open, disabled, pairing)
 * - Group policy enforcement (allowlist, open, disabled)
 * - Mention requirement in group conversations
 * - Message logging via ctx.logMessage
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockContext } from "./mocks/wopr-context.js";

// Track adapter.process calls to capture the turn handler
let capturedTurnHandler: ((context: any) => Promise<void>) | null = null;
const mockSendActivity = vi.fn().mockResolvedValue({});
const mockProcess = vi.fn(async (req: any, res: any, handler: any) => {
  capturedTurnHandler = handler;
  // Simulate the adapter calling the handler with a mock turn context
  if (req.__activity) {
    await handler({
      activity: req.__activity,
      sendActivity: mockSendActivity,
    });
  }
});

// Mock botbuilder
vi.mock("botbuilder", () => {
  return {
    CloudAdapter: class MockCloudAdapter {
      onTurnError: any;
      constructor() {
        this.onTurnError = null;
      }
      process = mockProcess;
      continueConversationAsync = vi.fn();
    },
    ConfigurationBotFrameworkAuthentication: class MockAuth {
      config: any;
      constructor(config: any) {
        this.config = config;
      }
    },
    TurnContext: class MockTurnContext {
      static getConversationReference(activity: any) {
        return {
          channelId: activity.channelId || "msteams",
          serviceUrl: activity.serviceUrl || "https://smba.trafficmanager.net/amer/",
          conversation: activity.conversation,
          bot: activity.recipient,
        };
      }
    },
    CardFactory: {
      adaptiveCard: vi.fn((card: any) => ({
        contentType: "application/vnd.microsoft.card.adaptive",
        content: card,
      })),
    },
    MessageFactory: {
      attachment: vi.fn((attachment: any) => ({
        type: "message",
        attachments: [attachment],
      })),
    },
  };
});

// Mock winston
vi.mock("winston", () => {
  const mockLogger = {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  };
  return {
    default: {
      createLogger: vi.fn(() => mockLogger),
      format: {
        combine: vi.fn(),
        timestamp: vi.fn(),
        errors: vi.fn(),
        json: vi.fn(),
        colorize: vi.fn(),
        simple: vi.fn(),
      },
      transports: {
        File: class MockFileTransport {
          constructor() {}
        },
        Console: class MockConsoleTransport {
          constructor() {}
        },
      },
    },
  };
});

// Mock axios
vi.mock("axios", () => ({
  default: {
    get: vi.fn(),
    post: vi.fn(),
  },
}));

function makeActivity(overrides: Record<string, any> = {}) {
  return {
    type: "message",
    text: "Hello WOPR",
    from: { id: "user-1", name: "Test User" },
    recipient: { id: "bot-1", name: "WOPR Bot" },
    conversation: {
      id: "conv-1",
      conversationType: "personal",
      name: "Test Chat",
    },
    ...overrides,
  };
}

describe("message handling", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    capturedTurnHandler = null;
    delete process.env.MSTEAMS_APP_ID;
    delete process.env.MSTEAMS_APP_PASSWORD;
    delete process.env.MSTEAMS_TENANT_ID;
  });

  afterEach(() => {
    vi.resetModules();
  });

  async function initPlugin(configData: Record<string, any> = {}) {
    const mod = await import("../src/index.js");
    const plugin = mod.default;

    const fullConfig = {
      appId: "test-app-id",
      appPassword: "test-password",
      tenantId: "test-tenant-id",
      useAdaptiveCards: false,
      ...configData,
    };

    const mockCtx = createMockContext(fullConfig);
    await plugin.init(mockCtx as any);
    return { mod, plugin, mockCtx };
  }

  it("handleWebhook delegates to adapter.process", async () => {
    const { mod } = await initPlugin();

    const mockReq = { __activity: makeActivity() };
    const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

    await mod.handleWebhook(mockReq, mockRes);

    expect(mockProcess).toHaveBeenCalledWith(
      mockReq,
      mockRes,
      expect.any(Function)
    );
  });

  it("processes valid DM and calls ctx.inject", async () => {
    const { mod, mockCtx } = await initPlugin({ dmPolicy: "open" });

    const activity = makeActivity({
      text: "Hello WOPR",
      from: { id: "user-1", name: "Alice" },
    });

    const mockReq = { __activity: activity };
    const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

    await mod.handleWebhook(mockReq, mockRes);

    expect(mockCtx.logMessage).toHaveBeenCalled();
    expect(mockCtx.inject).toHaveBeenCalledWith(
      "msteams-conv-1",
      "[Alice]: Hello WOPR",
      expect.objectContaining({
        from: "Alice",
        channel: expect.objectContaining({
          type: "msteams",
        }),
      })
    );
  });

  it("sends response back via sendActivity", async () => {
    const { mod, mockCtx } = await initPlugin({ dmPolicy: "open" });
    mockCtx.inject.mockResolvedValue("Bot says hello");

    const activity = makeActivity();
    const mockReq = { __activity: activity };
    const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

    await mod.handleWebhook(mockReq, mockRes);

    expect(mockSendActivity).toHaveBeenCalled();
  });

  it("skips non-message activities", async () => {
    const { mod, mockCtx } = await initPlugin({ dmPolicy: "open" });

    const activity = makeActivity({ type: "typing" });
    const mockReq = { __activity: activity };
    const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

    await mod.handleWebhook(mockReq, mockRes);

    expect(mockCtx.inject).not.toHaveBeenCalled();
  });

  it("skips messages from the bot itself", async () => {
    const { mod, mockCtx } = await initPlugin({ dmPolicy: "open" });

    const activity = makeActivity({
      from: { id: "bot-1", name: "WOPR Bot" },
      recipient: { id: "bot-1", name: "WOPR Bot" },
    });
    const mockReq = { __activity: activity };
    const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

    await mod.handleWebhook(mockReq, mockRes);

    expect(mockCtx.inject).not.toHaveBeenCalled();
  });

  it("skips messages with no userId", async () => {
    const { mod, mockCtx } = await initPlugin({ dmPolicy: "open" });

    const activity = makeActivity({
      from: { id: undefined, name: "Unknown" },
    });
    const mockReq = { __activity: activity };
    const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

    await mod.handleWebhook(mockReq, mockRes);

    expect(mockCtx.inject).not.toHaveBeenCalled();
  });

  it("skips messages with no conversationId", async () => {
    const { mod, mockCtx } = await initPlugin({ dmPolicy: "open" });

    const activity = makeActivity({
      conversation: { id: undefined, conversationType: "personal" },
    });
    const mockReq = { __activity: activity };
    const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

    await mod.handleWebhook(mockReq, mockRes);

    expect(mockCtx.inject).not.toHaveBeenCalled();
  });

  describe("DM policy", () => {
    it("open policy allows all DMs", async () => {
      const { mod, mockCtx } = await initPlugin({ dmPolicy: "open" });

      const activity = makeActivity();
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).toHaveBeenCalled();
    });

    it("disabled policy blocks all DMs", async () => {
      const { mod, mockCtx } = await initPlugin({ dmPolicy: "disabled" });

      const activity = makeActivity();
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).not.toHaveBeenCalled();
    });

    it("pairing policy allows all DMs", async () => {
      const { mod, mockCtx } = await initPlugin({ dmPolicy: "pairing" });

      const activity = makeActivity();
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).toHaveBeenCalled();
    });

    it("allowlist policy blocks users not in list", async () => {
      const { mod, mockCtx } = await initPlugin({
        dmPolicy: "allowlist",
        allowFrom: ["user-2", "user-3"],
      });

      const activity = makeActivity({
        from: { id: "user-1", name: "Blocked User" },
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).not.toHaveBeenCalled();
    });

    it("allowlist policy allows users in list", async () => {
      const { mod, mockCtx } = await initPlugin({
        dmPolicy: "allowlist",
        allowFrom: ["user-1", "user-2"],
      });

      const activity = makeActivity({
        from: { id: "user-1", name: "Allowed User" },
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).toHaveBeenCalled();
    });

    it("allowlist with wildcard allows all", async () => {
      const { mod, mockCtx } = await initPlugin({
        dmPolicy: "allowlist",
        allowFrom: ["*"],
      });

      const activity = makeActivity({
        from: { id: "anyone", name: "Anyone" },
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).toHaveBeenCalled();
    });
  });

  describe("group policy", () => {
    function makeGroupActivity(overrides: Record<string, any> = {}) {
      return makeActivity({
        conversation: {
          id: "group-1",
          conversationType: "channel",
          name: "General",
        },
        entities: [
          {
            type: "mention",
            mentioned: { id: "bot-1" },
          },
        ],
        ...overrides,
      });
    }

    it("open policy allows all group messages", async () => {
      const { mod, mockCtx } = await initPlugin({
        groupPolicy: "open",
        requireMention: false,
      });

      const activity = makeGroupActivity();
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).toHaveBeenCalled();
    });

    it("disabled policy blocks all group messages", async () => {
      const { mod, mockCtx } = await initPlugin({
        groupPolicy: "disabled",
      });

      const activity = makeGroupActivity();
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).not.toHaveBeenCalled();
    });

    it("allowlist blocks users not in group allow list", async () => {
      const { mod, mockCtx } = await initPlugin({
        groupPolicy: "allowlist",
        groupAllowFrom: ["user-5"],
        requireMention: false,
      });

      const activity = makeGroupActivity({
        from: { id: "user-1", name: "Blocked" },
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).not.toHaveBeenCalled();
    });

    it("allowlist falls back to allowFrom when groupAllowFrom not set", async () => {
      const { mod, mockCtx } = await initPlugin({
        groupPolicy: "allowlist",
        allowFrom: ["user-1"],
        requireMention: false,
      });

      const activity = makeGroupActivity({
        from: { id: "user-1", name: "Allowed" },
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).toHaveBeenCalled();
    });
  });

  describe("mention requirement", () => {
    it("requires mention in channel by default", async () => {
      const { mod, mockCtx } = await initPlugin({
        groupPolicy: "open",
      });

      const activity = makeActivity({
        conversation: {
          id: "group-1",
          conversationType: "channel",
          name: "General",
        },
        entities: [], // no mention
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).not.toHaveBeenCalled();
    });

    it("processes message when bot is mentioned in channel", async () => {
      const { mod, mockCtx } = await initPlugin({
        groupPolicy: "open",
      });

      const activity = makeActivity({
        conversation: {
          id: "group-1",
          conversationType: "channel",
          name: "General",
        },
        entities: [
          {
            type: "mention",
            mentioned: { id: "bot-1" },
          },
        ],
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).toHaveBeenCalled();
    });

    it("skips mention check when requireMention is false", async () => {
      const { mod, mockCtx } = await initPlugin({
        groupPolicy: "open",
        requireMention: false,
      });

      const activity = makeActivity({
        conversation: {
          id: "group-1",
          conversationType: "channel",
          name: "General",
        },
        entities: [],
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).toHaveBeenCalled();
    });

    it("applies mention check in groupChat conversations too", async () => {
      const { mod, mockCtx } = await initPlugin({
        groupPolicy: "open",
      });

      const activity = makeActivity({
        conversation: {
          id: "group-2",
          conversationType: "groupChat",
          name: "Team Chat",
        },
        entities: [],
      });
      const mockReq = { __activity: activity };
      const mockRes = { status: vi.fn().mockReturnThis(), send: vi.fn() };

      await mod.handleWebhook(mockReq, mockRes);

      expect(mockCtx.inject).not.toHaveBeenCalled();
    });
  });
});
