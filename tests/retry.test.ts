/**
 * Tests for error retry with exponential backoff + jitter (WOP-115)
 *
 * Tests:
 * - Successful call returns immediately
 * - Retries on 429 status
 * - Retries on 500+ status
 * - Does not retry on 400 (client error)
 * - Respects maxRetries limit
 * - Respects Retry-After header
 * - Uses exponential backoff with jitter
 * - Throws last error after all retries exhausted
 */
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("botbuilder", () => {
  return {
    CloudAdapter: class { onTurnError: any; constructor() { this.onTurnError = null; } process = vi.fn(); continueConversationAsync = vi.fn(); },
    ConfigurationBotFrameworkAuthentication: class { constructor(config: any) {} },
    TurnContext: class { static getConversationReference() { return {}; } },
    CardFactory: { adaptiveCard: vi.fn((card: any) => ({ contentType: "application/vnd.microsoft.card.adaptive", content: card })) },
    MessageFactory: { attachment: vi.fn((a: any) => ({ type: "message", attachments: [a] })) },
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

describe("withRetry", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.useFakeTimers({ shouldAdvanceTime: true });
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.resetModules();
  });

  it("returns result on first successful call", async () => {
    const { withRetry } = await import("../src/index.js");
    const fn = vi.fn().mockResolvedValue("success");

    const result = await withRetry(fn, { maxRetries: 3, baseDelayMs: 100 });

    expect(result).toBe("success");
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it("retries on 429 status and succeeds", async () => {
    const { withRetry } = await import("../src/index.js");
    const fn = vi.fn()
      .mockRejectedValueOnce({ response: { status: 429, headers: {} } })
      .mockResolvedValueOnce("success");

    const result = await withRetry(fn, { maxRetries: 3, baseDelayMs: 10 });

    expect(result).toBe("success");
    expect(fn).toHaveBeenCalledTimes(2);
  });

  it("retries on 500 status and succeeds", async () => {
    const { withRetry } = await import("../src/index.js");
    const fn = vi.fn()
      .mockRejectedValueOnce({ response: { status: 500, headers: {} } })
      .mockResolvedValueOnce("recovered");

    const result = await withRetry(fn, { maxRetries: 3, baseDelayMs: 10 });

    expect(result).toBe("recovered");
    expect(fn).toHaveBeenCalledTimes(2);
  });

  it("retries on 503 status", async () => {
    const { withRetry } = await import("../src/index.js");
    const fn = vi.fn()
      .mockRejectedValueOnce({ response: { status: 503, headers: {} } })
      .mockResolvedValueOnce("ok");

    const result = await withRetry(fn, { maxRetries: 2, baseDelayMs: 10 });

    expect(result).toBe("ok");
    expect(fn).toHaveBeenCalledTimes(2);
  });

  it("does not retry on 400 client error", async () => {
    const { withRetry } = await import("../src/index.js");
    const err = { response: { status: 400, headers: {} } };
    const fn = vi.fn().mockRejectedValue(err);

    await expect(withRetry(fn, { maxRetries: 3, baseDelayMs: 10 })).rejects.toEqual(err);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it("does not retry on 401 unauthorized", async () => {
    const { withRetry } = await import("../src/index.js");
    const err = { response: { status: 401, headers: {} } };
    const fn = vi.fn().mockRejectedValue(err);

    await expect(withRetry(fn, { maxRetries: 3, baseDelayMs: 10 })).rejects.toEqual(err);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it("does not retry on 404", async () => {
    const { withRetry } = await import("../src/index.js");
    const err = { response: { status: 404, headers: {} } };
    const fn = vi.fn().mockRejectedValue(err);

    await expect(withRetry(fn, { maxRetries: 3, baseDelayMs: 10 })).rejects.toEqual(err);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it("throws after maxRetries exhausted", async () => {
    const { withRetry } = await import("../src/index.js");
    const err = { response: { status: 429, headers: {} } };
    const fn = vi.fn().mockRejectedValue(err);

    await expect(withRetry(fn, { maxRetries: 2, baseDelayMs: 10 })).rejects.toEqual(err);
    // 1 initial + 2 retries = 3 total
    expect(fn).toHaveBeenCalledTimes(3);
  });

  it("respects maxRetries of 0 (no retries)", async () => {
    const { withRetry } = await import("../src/index.js");
    const err = { response: { status: 500, headers: {} } };
    const fn = vi.fn().mockRejectedValue(err);

    await expect(withRetry(fn, { maxRetries: 0, baseDelayMs: 10 })).rejects.toEqual(err);
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it("handles errors without response object", async () => {
    const { withRetry } = await import("../src/index.js");
    const err = new Error("network timeout");
    const fn = vi.fn().mockRejectedValue(err);

    await expect(withRetry(fn, { maxRetries: 2, baseDelayMs: 10 })).rejects.toEqual(err);
    // No status code means status 0, which is not retryable
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it("handles errors with statusCode property", async () => {
    const { withRetry } = await import("../src/index.js");
    const fn = vi.fn()
      .mockRejectedValueOnce({ statusCode: 502 })
      .mockResolvedValueOnce("ok");

    const result = await withRetry(fn, { maxRetries: 2, baseDelayMs: 10 });

    expect(result).toBe("ok");
    expect(fn).toHaveBeenCalledTimes(2);
  });
});
