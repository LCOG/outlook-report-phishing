import { vi } from "vitest";

export function createSuccessResult<T>(value: T): Office.AsyncResult<T> {
  return {
    status: Office.AsyncResultStatus.Succeeded,
    value,
    error: null as unknown as Office.Error,
  } as Office.AsyncResult<T>;
}

export function createFailureResult<T>(message: string): Office.AsyncResult<T> {
  return {
    status: Office.AsyncResultStatus.Failed,
    value: null as unknown as T,
    error: {
      name: "Error",
      message,
      code: 9001,
    },
  } as Office.AsyncResult<T>;
}

export function mockOfficeGlobal(): void {
  // @ts-ignore
  global.Office = {
    context: {
      mailbox: {
        item: {
          subject: { getAsync: vi.fn() },
          from: { getAsync: vi.fn() },
          to: { getAsync: vi.fn() },
          cc: { getAsync: vi.fn() },
          itemId: "test-item-id",
          conversationId: "conv-id",
          dateTimeCreated: new Date("2024-01-01T12:00:00Z"),
          dateTimeModified: new Date("2024-01-01T12:00:00Z"),
          internetMessageId: "msg-id",
          notificationMessages: {
            addAsync: vi.fn(),
          },
          getAttachmentsAsync: vi.fn(),
          internetHeaders: {
            getAsync: vi.fn(),
          },
        },
        diagnostics: {
          hostName: "Outlook",
        },
        convertToRestId: vi.fn(),
      },
    },
    AsyncResultStatus: {
      Succeeded: "succeeded",
      Failed: "failed",
    },
    MailboxEnums: {
      RestVersion: {
        v2_0: "v2.0",
      },
      ItemNotificationMessageType: {
        InformationalMessage: "informationalMessage",
      },
    },
    onReady: vi.fn(async () => Promise.resolve()),
  } as unknown as typeof Office;
}
