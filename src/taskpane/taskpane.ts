/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console, process, fetch */

import {
  provideFluentDesignSystem,
  fluentButton,
  fluentDivider,
  fluentRadio,
  fluentRadioGroup,
  fluentTextArea,
  fluentTooltip,
  fluentProgressRing,
} from "@fluentui/web-components";

import { AccountManager } from "./authConfig";
import { makeGraphRequest, makePostGraphRequest } from "./msgraph-helper";

provideFluentDesignSystem().register(
  fluentButton(),
  fluentDivider(),
  fluentRadio(),
  fluentRadioGroup(),
  fluentTextArea(),
  fluentTooltip(),
  fluentProgressRing()
);

// Type definitions
enum UIState {
  IDLE = "idle",
  REPORTING = "reporting",
  SUCCESS = "success",
  ERROR = "error",
}

interface UIElements {
  reportEmailButton: HTMLElement | null;
  progress: HTMLElement | null;
  success: HTMLElement | null;
  error: HTMLElement | null;
  errorMessage: HTMLElement | null;
  moveToJunkButton: HTMLElement | null;
  reportMessageTypeGroup: any;
  additionalInfoArea: any;
  sideloadMsg: HTMLElement | null;
  appBody: HTMLElement | null;
}

// State management
let currentState: UIState = UIState.IDLE;

// Cache DOM elements
const elements: UIElements = {
  reportEmailButton: null,
  progress: null,
  success: null,
  error: null,
  errorMessage: null,
  moveToJunkButton: null,
  reportMessageTypeGroup: null,
  additionalInfoArea: null,
  sideloadMsg: null,
  appBody: null,
};

const reportPhishApiUrl = process.env.API_URL;
const forwardEmail = process.env.FORWARD_TO;
const accountManager = new AccountManager();

function initializeElements(): void {
  // Cache all DOM element references with proper type assertions
  elements.reportEmailButton = document.getElementById("reportEmailButton");
  elements.progress = document.getElementById("reportingProgress");
  elements.success = document.getElementById("reportingSuccess");
  elements.error = document.getElementById("reportingError");
  elements.errorMessage = document.getElementById("reportingErrorMessage");
  elements.moveToJunkButton = document.getElementById("moveToJunkButton");
  elements.reportMessageTypeGroup = document.getElementById("reportMessageTypeGroup");
  elements.additionalInfoArea = document.getElementById("additionalInfo");
  elements.sideloadMsg = document.getElementById("sideload-msg");
  elements.appBody = document.getElementById("app-body");

  // Verify critical elements exist
  if (
    !elements.reportEmailButton ||
    !elements.progress ||
    !elements.success ||
    !elements.error ||
    !elements.errorMessage
  ) {
    console.error("Required DOM elements not found");
  }
}

/**
 * Updates the UI based on the provided state
 * @param state The new UI state
 * @param errorMessage Optional error message for ERROR state
 */
function updateUIState(state: UIState, errorMessage: string = ""): void {
  currentState = state;

  // Hides all reporting result elements
  const results = [elements.progress, elements.success, elements.error];
  results.forEach((el) => el?.classList.add("hidden"));

  // Reset button state
  if (elements.reportEmailButton) {
    (elements.reportEmailButton as any).disabled = state === UIState.REPORTING;
  }

  // Show the appropriate element based on the current state
  switch (state) {
    case UIState.REPORTING:
      elements.progress?.classList.remove("hidden");
      break;
    case UIState.SUCCESS:
      elements.success?.classList.remove("hidden");
      elements.moveToJunkButton?.classList.remove("hidden");
      break;
    case UIState.ERROR:
      elements.error?.classList.remove("hidden");
      if (elements.errorMessage) {
        elements.errorMessage.innerText = errorMessage || "Something went wrong";
      }
      break;
    case UIState.IDLE:
      // Everything hidden is the default idle state
      break;
  }
}

// Initialize when Office is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initializeElements();

    if (elements.sideloadMsg) elements.sideloadMsg.style.display = "none";
    if (elements.appBody) elements.appBody.style.display = "flex";

    if (elements.reportEmailButton) {
      elements.reportEmailButton.addEventListener("click", handleReportClick);
    }
    if (elements.moveToJunkButton) {
      elements.moveToJunkButton.addEventListener("click", handleMoveToJunkClick);
    }

    // Initialize MSAL.
    accountManager.initialize();
    updateUIState(UIState.IDLE);
  }
});

/**
 * Gets the user's name and email
 */
async function getUserData() {
  const accessToken = await accountManager.ssoGetAccessToken(["user.read"]);
  const response: { displayName: string; mail: string } = await makeGraphRequest({
    accessToken,
    path: "/me",
  });
  return response;
}

async function logPhishingReport(emailAddress: string, message: any) {
  try {
    const response = await fetch(reportPhishApiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        employee_email: emailAddress.toLowerCase(),
        email_message: message.body.content,
      }),
    });

    const result = await response.json();
    console.log("Phishing report logged:", result);
  } catch (error) {
    console.error("Failed to log phishing report:", error.message);
  }
}

async function handleReportClick(): Promise<void> {
  try {
    updateUIState(UIState.REPORTING);

    const accessToken: string = await accountManager.ssoGetAccessToken(["mail.read", "mail.send"]);

    // The type of Office.context.mailbox.item can vary depending on the context
    // so we explicitly assign the expected type here.
    // TODO: Note that this is a suspected phishing email, and any links and
    // attachments on it need to be treated as malicious.
    // Research best practices on handling such emails.
    const msg: Office.MessageRead = Office.context.mailbox.item as Office.MessageRead;

    if (!msg.itemId) {
      throw new Error("Failed to retrieve current email from Office context.");
    }

    const msgId = Office.context.mailbox.convertToRestId(
      msg.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
    const { displayName, mail } = await getUserData();

    // Get the current email subject and body via Graph API
    const currentMessageBody = await makeGraphRequest({
      accessToken: accessToken,
      path: `/me/messages/${msgId}`,
      queryParams: "?$select=subject,body",
      additionalHeaders: { Prefer: 'outlook.body-content-type="text"' },
    });

    await logPhishingReport(mail, currentMessageBody);

    // Set body for the forwarding request
    const reportType = elements.reportMessageTypeGroup?.value || "unknown";
    const additionalInfo = elements.additionalInfoArea?.value || "";
    const forwardComment = `${displayName} forwarded a suspicious email (${reportType}) via the Report Phish add-in. ${additionalInfo ? `\n\nAdditional details: ${additionalInfo}` : ""}`;

    const forwardBody = {
      comment: forwardComment,
      toRecipients: [
        {
          emailAddress: {
            name: "LCOG IT",
            address: forwardEmail,
          },
        },
      ],
    };

    const forwardResult = await makePostGraphRequest({
      accessToken,
      path: `/me/messages/${msgId}/forward`,
      additionalHeaders: { "Content-Type": "application/json" },
      body: JSON.stringify(forwardBody),
    });

    if (forwardResult.ok) {
      updateUIState(UIState.SUCCESS);
      // Auto-dismiss success message after 3 seconds
      setTimeout(() => {
        if (currentState === UIState.SUCCESS) {
          updateUIState(UIState.IDLE);
        }
      }, 3000);
    } else {
      throw new Error(`HTTP ${forwardResult.status} ${forwardResult.statusText}`);
    }
  } catch (error) {
    console.error("Report clicking failed:", error);
    updateUIState(
      UIState.ERROR,
      error instanceof Error ? error.message : "Reporting the email failed"
    );
  }
}

async function handleMoveToJunkClick(): Promise<void> {
  try {
    updateUIState(UIState.REPORTING);
    const accessToken: string = await accountManager.ssoGetAccessToken(["mail.read", "mail.send"]);
    const msg: Office.MessageRead = Office.context.mailbox.item as Office.MessageRead;

    if (!msg.itemId) {
      throw new Error("Failed to retrieve current email from Office context.");
    }

    const msgId = Office.context.mailbox.convertToRestId(
      msg.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    const response = await makePostGraphRequest({
      accessToken,
      path: `/me/messages/${msgId}/move`,
      additionalHeaders: { "Content-Type": "application/json" },
      body: JSON.stringify({ destinationId: "junkemail" }),
    });

    if (response.ok) {
      Office.context.ui.closeContainer();
    } else {
      throw new Error(`${response.status} ${response.statusText}`);
    }
  } catch (error) {
    console.error("Moving to junk failed:", error);
    updateUIState(UIState.ERROR, error instanceof Error ? error.message : "Failed to move email");
  }
}
