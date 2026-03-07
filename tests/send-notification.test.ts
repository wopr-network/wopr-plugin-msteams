/**
 * Tests for sendNotification() on the msteams channel provider (WOP-1668)
 *
 * Tests:
 * - sendNotification is present on the registered provider
 * - Ignores non-friend-request notification types
 * - No-ops when no conversation reference exists
 * - Sends Adaptive Card via continueConversationAsync for friend-request,
 *   asserting card body text, Action.Submit buttons, and channelId payload
 * - Normalizes channelId without "msteams:" prefix (both lookup paths)
 * - Stores callbacks keyed on activity ID and fires onAccept/onDeny on invoke
 * - Always sends invokeResponse 200 even when no matching callback exists
 * - Clears pendingCallbacks on shutdown so callbacks no longer fire
 */

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockContext } from "./mocks/wopr-context.js";

const mockContinueConversation = vi.fn();
const mockProcess = vi.fn();

vi.mock("botbuilder", () => {
  return {
    CloudAdapter: class MockCloudAdapter {
      onTurnError: any;
      constructor() {
        this.onTurnError = null;
      }
      process = mockProcess;
      continueConversationAsync = mockContinueConversation;
    },
    ConfigurationBotFrameworkAuthentication: class MockAuth {
      constructor(_config: any) {}
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
    MessageFactory: {
      attachment: (att: any) => ({ attachments: [att], type: "message" }),
    },
    CardFactory: {
      adaptiveCard: (card: any) => ({
        contentType: "application/vnd.microsoft.card.adaptive",
        content: card,
      }),
    },
  };
});

vi.mock("axios", () => ({
  default: { get: vi.fn(), post: vi.fn() },
}));

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
        combine: vi.fn(() => ({})),
        timestamp: vi.fn(() => ({})),
        errors: vi.fn(() => ({})),
        json: vi.fn(() => ({})),
        colorize: vi.fn(() => ({})),
        simple: vi.fn(() => ({})),
      },
      transports: {
        File: class MockFile {
          constructor(_opts: any) {}
        },
        Console: class MockConsole {
          constructor(_opts: any) {}
        },
      },
    },
  };
});

describe("sendNotification", () => {
  let plugin: any;
  let handleWebhook: any;
  let mockCtx: any;

  beforeEach(async () => {
    vi.resetModules();
    vi.clearAllMocks();

    // Default: continueConversationAsync calls the callback with a mock turn context
    mockContinueConversation.mockImplementation(async (_appId: string, _ref: any, callback: Function) => {
      const mockTurnContext = {
        sendActivity: vi.fn().mockResolvedValue({ id: "activity-123" }),
        sendInvokeResponse: vi.fn().mockResolvedValue(undefined),
      };
      await callback(mockTurnContext);
    });

    const module = await import("../src/index.js");
    plugin = module.default;
    handleWebhook = module.handleWebhook;
    mockCtx = createMockContext({
      appId: "test-app-id",
      appPassword: "test-password",
      tenantId: "test-tenant",
    });
    await plugin.init(mockCtx);
  });

  afterEach(async () => {
    await plugin.shutdown();
  });

  function getProvider() {
    const calls = mockCtx.registerChannelProvider.mock.calls;
    if (!calls.length) throw new Error("registerChannelProvider was not called");
    return calls[0][0];
  }

  /**
   * Drive a fake inbound message through handleWebhook → processActivity
   * so that a conversation reference is stored for the given convId.
   */
  async function storeConvRef(convId: string): Promise<void> {
    let processCallback: Function | null = null;
    mockProcess.mockImplementationOnce(async (_req: any, _res: any, callback: Function) => {
      processCallback = callback;
    });

    const activity = {
      type: "message",
      text: "hello",
      from: { id: "user-1", name: "Alice" },
      conversation: { id: convId, isGroup: false },
      recipient: { id: "bot-1", name: "WOPR" },
      channelId: "msteams",
      serviceUrl: "https://smba.trafficmanager.net/amer/",
    };

    await handleWebhook({ body: activity, headers: {} }, { status: vi.fn().mockReturnThis(), end: vi.fn() });

    if (processCallback) {
      const mockTurn = {
        activity,
        sendActivity: vi.fn().mockResolvedValue({ id: "ref-msg" }),
      };
      await (processCallback as Function)(mockTurn);
    }
  }

  /**
   * Drive an invoke activity through handleWebhook → processActivity.
   * Returns the sendActivity mock so callers can assert on it.
   */
  async function driveInvoke(invokeActivity: any): Promise<ReturnType<typeof vi.fn>> {
    let processCallback: Function | null = null;
    mockProcess.mockImplementationOnce(async (_req: any, _res: any, callback: Function) => {
      processCallback = callback;
    });

    await handleWebhook({ body: invokeActivity, headers: {} }, { status: vi.fn().mockReturnThis(), end: vi.fn() });

    const sendActivity = vi.fn().mockResolvedValue({ id: "invoke-resp" });
    if (processCallback) {
      await (processCallback as Function)({ activity: invokeActivity, sendActivity });
    }
    return sendActivity;
  }

  it("should have sendNotification on the channel provider", () => {
    const provider = getProvider();
    expect(typeof provider.sendNotification).toBe("function");
  });

  it("should ignore non-friend-request notification types", async () => {
    const provider = getProvider();

    await provider.sendNotification("msteams:conv-1", { type: "unknown-type" }, { onAccept: vi.fn(), onDeny: vi.fn() });

    expect(mockContinueConversation).not.toHaveBeenCalled();
  });

  it("should no-op when no conversation reference exists for channelId", async () => {
    const provider = getProvider();

    await provider.sendNotification(
      "msteams:nonexistent-conv",
      { type: "friend-request", from: "Alice" },
      { onAccept: vi.fn(), onDeny: vi.fn() },
    );

    expect(mockContinueConversation).not.toHaveBeenCalled();
  });

  it("should send an adaptive card via continueConversationAsync for friend-request", async () => {
    const provider = getProvider();

    await storeConvRef("conv-1");

    // Override to capture what gets sent to sendActivity
    let capturedSendActivity: ReturnType<typeof vi.fn> | null = null;
    mockContinueConversation.mockImplementationOnce(async (_appId: string, _ref: any, callback: Function) => {
      const sendActivity = vi.fn().mockResolvedValue({ id: "activity-123" });
      capturedSendActivity = sendActivity;
      await callback({ sendActivity });
    });

    await provider.sendNotification(
      "msteams:conv-1",
      { type: "friend-request", from: "Alice" },
      { onAccept: vi.fn(), onDeny: vi.fn() },
    );

    expect(mockContinueConversation).toHaveBeenCalled();
    const lastCall = mockContinueConversation.mock.calls[mockContinueConversation.mock.calls.length - 1];
    expect(lastCall[0]).toBe("test-app-id");

    // Verify the Adaptive Card was sent with correct content
    expect(capturedSendActivity).not.toBeNull();
    expect(capturedSendActivity).toHaveBeenCalledTimes(1);
    const sentActivity = capturedSendActivity!.mock.calls[0][0];
    expect(sentActivity.attachments?.[0]?.contentType).toBe("application/vnd.microsoft.card.adaptive");

    const card = sentActivity.attachments[0].content;
    const bodyTexts = (card.body ?? [])
      .map((b: any) => b?.text)
      .filter((t: unknown): t is string => typeof t === "string");
    expect(bodyTexts.join(" ")).toContain("Alice");
    expect(bodyTexts.join(" ").toLowerCase()).toContain("connect");

    // Verify Action.Submit buttons with correct action names and channelId
    const submitActions = (card.actions ?? []).filter((a: any) => a?.type === "Action.Submit");
    expect(submitActions).toHaveLength(2);

    const accept = submitActions.find((a: any) => a?.data?.action === "friend-request-accept");
    const deny = submitActions.find((a: any) => a?.data?.action === "friend-request-deny");
    expect(accept).toBeDefined();
    expect(deny).toBeDefined();
    expect(accept!.data.channelId).toBe("msteams:conv-1");
    expect(deny!.data.channelId).toBe("msteams:conv-1");
  });

  it("should find conversation reference when channelId has no msteams: prefix", async () => {
    const provider = getProvider();

    await storeConvRef("conv-2");

    let capturedSendActivity: ReturnType<typeof vi.fn> | null = null;
    mockContinueConversation.mockImplementationOnce(async (_appId: string, _ref: any, callback: Function) => {
      const sendActivity = vi.fn().mockResolvedValue({ id: "activity-456" });
      capturedSendActivity = sendActivity;
      await callback({ sendActivity });
    });

    // Pass the bare convId without the "msteams:" prefix
    await provider.sendNotification(
      "conv-2",
      { type: "friend-request", from: "Bob" },
      { onAccept: vi.fn(), onDeny: vi.fn() },
    );

    // Should still call continueConversationAsync (reference lookup succeeded)
    expect(mockContinueConversation).toHaveBeenCalled();
    expect(capturedSendActivity).not.toBeNull();

    // channelId in action data reflects what was passed in (no forced normalization)
    const card = capturedSendActivity!.mock.calls[0][0].attachments[0].content;
    const accept = (card.actions ?? []).find(
      (a: any) => a?.type === "Action.Submit" && a?.data?.action === "friend-request-accept",
    );
    expect(accept).toBeDefined();
    expect(accept!.data.channelId).toBe("conv-2");
  });

  it("should fire onAccept callback when invoke activity with friend-request-accept arrives", async () => {
    const provider = getProvider();

    await storeConvRef("conv-1");

    const onAccept = vi.fn().mockResolvedValue(undefined);
    const onDeny = vi.fn().mockResolvedValue(undefined);

    // sendNotification stores callbacks under the activity ID returned by sendActivity
    // The default beforeEach mock returns { id: "activity-123" }
    await provider.sendNotification("msteams:conv-1", { type: "friend-request", from: "Alice" }, { onAccept, onDeny });

    // Drive an invoke with replyToId matching the stored activity ID
    const invokeSendActivity = await driveInvoke({
      type: "invoke",
      replyToId: "activity-123",
      value: { action: "friend-request-accept", channelId: "msteams:conv-1" },
      from: { id: "user-1" },
      conversation: { id: "conv-1" },
      channelId: "msteams",
      serviceUrl: "https://smba.trafficmanager.net/amer/",
    });

    expect(onAccept).toHaveBeenCalledTimes(1);
    expect(onDeny).not.toHaveBeenCalled();

    // Verify 200 invokeResponse was sent
    expect(invokeSendActivity).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "invokeResponse",
        value: { status: 200, body: {} },
      }),
    );
  });

  it("should fire onDeny callback when invoke activity with friend-request-deny arrives", async () => {
    const provider = getProvider();

    await storeConvRef("conv-1");

    const onAccept = vi.fn().mockResolvedValue(undefined);
    const onDeny = vi.fn().mockResolvedValue(undefined);

    await provider.sendNotification("msteams:conv-1", { type: "friend-request", from: "Alice" }, { onAccept, onDeny });

    const invokeSendActivity = await driveInvoke({
      type: "invoke",
      replyToId: "activity-123",
      value: { action: "friend-request-deny", channelId: "msteams:conv-1" },
      from: { id: "user-1" },
      conversation: { id: "conv-1" },
      channelId: "msteams",
      serviceUrl: "https://smba.trafficmanager.net/amer/",
    });

    expect(onDeny).toHaveBeenCalledTimes(1);
    expect(onAccept).not.toHaveBeenCalled();
    expect(invokeSendActivity).toHaveBeenCalledWith(
      expect.objectContaining({ type: "invokeResponse", value: { status: 200, body: {} } }),
    );
  });

  it("should always send invokeResponse 200 even when no matching callback exists", async () => {
    const invokeSendActivity = await driveInvoke({
      type: "invoke",
      replyToId: "no-such-activity-id",
      value: { action: "friend-request-accept", channelId: "msteams:conv-1" },
      from: { id: "user-1" },
      conversation: { id: "conv-1" },
      channelId: "msteams",
      serviceUrl: "https://smba.trafficmanager.net/amer/",
    });

    expect(invokeSendActivity).toHaveBeenCalledWith(
      expect.objectContaining({
        type: "invokeResponse",
        value: { status: 200, body: {} },
      }),
    );
  });

  it("should clear pendingCallbacks on shutdown so callbacks no longer fire", async () => {
    const provider = getProvider();

    await storeConvRef("conv-1");

    const onAccept = vi.fn().mockResolvedValue(undefined);

    // Register callbacks
    await provider.sendNotification(
      "msteams:conv-1",
      { type: "friend-request", from: "Alice" },
      { onAccept, onDeny: vi.fn() },
    );

    // Shutdown clears pendingCallbacks
    await plugin.shutdown();

    // Re-init so the adapter is available for handleWebhook
    await plugin.init(mockCtx);

    // Drive invoke with the previously registered activity ID — callbacks are gone
    const invokeSendActivity = await driveInvoke({
      type: "invoke",
      replyToId: "activity-123",
      value: { action: "friend-request-accept", channelId: "msteams:conv-1" },
      from: { id: "user-1" },
      conversation: { id: "conv-1" },
      channelId: "msteams",
      serviceUrl: "https://smba.trafficmanager.net/amer/",
    });

    // Callback must not fire — it was cleared by shutdown
    expect(onAccept).not.toHaveBeenCalled();

    // But the 200 response is still sent (invoke always gets a response)
    expect(invokeSendActivity).toHaveBeenCalledWith(
      expect.objectContaining({ type: "invokeResponse", value: { status: 200, body: {} } }),
    );
  });
});
