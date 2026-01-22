import { vi } from "vitest";

// Mock Office.js global
const OfficeMock = {
  context: {
    mailbox: {
      item: {
        subject: {
          getAsync: vi.fn(),
        },
        from: {
          getAsync: vi.fn(),
        },
        to: {
          getAsync: vi.fn(),
        },
      },
    },
  },
  onReady: vi.fn(async () => Promise.resolve()),
  EventType: {
    DialogMessageReceived: "dialogMessageReceived",
    DialogEventReceived: "dialogEventReceived",
  },
  MailboxEnums: {
    RestVersion: {
      v2_0: "v2.0",
    },
  },
};
(globalThis as unknown as { Office: typeof OfficeMock }).Office = OfficeMock;
(globalThis as unknown as { OfficeExtension: object }).OfficeExtension = {};
