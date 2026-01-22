/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

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
import { GraphClient } from "../services/graph-client";

import type { Message, User } from "@microsoft/microsoft-graph-types";

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
  reportEmailTooltip: HTMLElement | null;
  reportInProgress: HTMLElement | null;
  reportSuccess: HTMLElement | null;
  reportError: HTMLElement | null;
  reportErrorMessage: HTMLElement | null;
  moveToJunkButton: HTMLElement | null;
  reportMessageTypeGroup: { value: string } | null;
  additionalInfoArea: { value: string } | null;
  sideloadMsg: HTMLElement | null;
  appBody: HTMLElement | null;
}

// Cache DOM elements
const elements: UIElements = {
  reportEmailButton: null,
  reportEmailTooltip: null,
  reportInProgress: null,
  reportSuccess: null,
  reportError: null,
  reportErrorMessage: null,
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
  elements.reportEmailTooltip = document.getElementById("reportEmailTooltip");
  elements.reportInProgress = document.getElementById("reportInProgress");
  elements.reportSuccess = document.getElementById("reportingSuccess");
  elements.reportError = document.getElementById("reportingError");
  elements.reportErrorMessage = document.getElementById("reportingErrorMessage");
  elements.moveToJunkButton = document.getElementById("moveToJunkButton");
  elements.reportMessageTypeGroup = document.getElementById(
    "reportMessageTypeGroup"
  ) as unknown as { value: string } | null;
  elements.additionalInfoArea = document.getElementById("additionalInfo") as unknown as {
    value: string;
  } | null;
  elements.sideloadMsg = document.getElementById("sideload-msg");
  elements.appBody = document.getElementById("app-body");

  // Verify critical elements exist
  if (
    !elements.reportEmailButton ||
    !elements.reportInProgress ||
    !elements.reportSuccess ||
    !elements.reportError ||
    !elements.reportErrorMessage ||
    !elements.moveToJunkButton ||
    !elements.reportMessageTypeGroup ||
    !elements.additionalInfoArea ||
    !elements.appBody
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
  // Hides all reporting result elements
  const results = [elements.reportInProgress, elements.reportSuccess, elements.reportError];
  results.forEach((el) => el?.classList.add("hidden"));

  // Show the appropriate element based on the current state
  switch (state) {
    case UIState.REPORTING:
      elements.reportEmailButton?.setAttribute("disabled", "true");
      if (elements.reportEmailTooltip) {
        elements.reportEmailTooltip.innerText = "Report in progress, please wait...";
      }
      elements.reportInProgress?.classList.remove("hidden");
      break;
    case UIState.SUCCESS:
      elements.reportEmailButton?.setAttribute("disabled", "true");
      if (elements.reportEmailTooltip) {
        elements.reportEmailTooltip.innerText = "Email reported successfully";
      }
      elements.reportSuccess?.classList.remove("hidden");
      elements.moveToJunkButton?.classList.remove("hidden");
      break;
    case UIState.ERROR:
      elements.reportEmailButton?.setAttribute("disabled", "false");
      if (elements.reportEmailTooltip) {
        elements.reportEmailTooltip.innerText = "Click to retry reporting the email";
      }
      elements.reportError?.classList.remove("hidden");
      if (elements.reportErrorMessage) {
        elements.reportErrorMessage.innerText = errorMessage || "Something went wrong";
      }
      break;
    case UIState.IDLE:
      // Everything hidden is the default idle state
      if (elements.reportEmailTooltip) {
        elements.reportEmailTooltip.innerText = "Report this email as a phishing or a spam message";
      }
      break;
  }
}

// Initialize when Office is ready.
void Office.onReady(async (info) => {
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
    await accountManager.initialize();
    updateUIState(UIState.IDLE);
  }
});

/**
 * Gets the user's name and email
 */
async function getUserData(): Promise<User> {
  const accessToken = await accountManager.ssoGetAccessToken(["user.read"]);
  const graphClient = new GraphClient(accessToken);
  return graphClient.getUser();
}

async function logPhishingReport(emailAddress: string, message: Message): Promise<void> {
  try {
    await fetch(reportPhishApiUrl as string, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        employee_email: emailAddress.toLowerCase(),
        email_message: message.body?.content ?? "No content",
      }),
    });
  } catch (error) {
    if (error instanceof Error) {
      console.error("Failed to log phishing report:", error.message);
    }
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
    const user = await getUserData();
    const mail = user.mail ?? "";
    const displayName = user.displayName ?? "User";

    const graphClient = new GraphClient(accessToken);

    // Get the current email subject and body via Graph API
    const currentMessageBody = await graphClient.getMessage(msgId, "?$select=subject,body");

    // Set Prefer header if needed, but our GraphClient doesn't support additional headers in getMessage yet
    // I will add it or use request<Message> directly if I need custom headers.
    // For now, let's just use request<Message> for this one to keep the Prefer header.
    // const currentMessageBody = await graphClient.get<Message>(`/me/messages/${msgId}?$select=subject,body`, {
    //   headers: { Prefer: 'outlook.body-content-type="text"' }
    // });

    await logPhishingReport(mail, currentMessageBody);

    // Set body for the forwarding request
    const reportType = elements.reportMessageTypeGroup?.value ?? "unknown";
    const additionalInfo = elements.additionalInfoArea?.value ?? "";
    const forwardComment = `${displayName} forwarded a suspicious email (${reportType}) via the Report Phish add-in. ${additionalInfo !== "" ? `\n\nAdditional details: ${additionalInfo}` : ""}`;

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

    await graphClient.forwardMessage(msgId, forwardBody);
    updateUIState(UIState.SUCCESS);
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
    const accessToken: string = await accountManager.ssoGetAccessToken(["mail.read", "mail.send"]);
    const msg: Office.MessageRead = Office.context.mailbox.item as Office.MessageRead;

    if (!msg.itemId) {
      throw new Error("Failed to retrieve current email from Office context.");
    }

    const msgId = Office.context.mailbox.convertToRestId(
      msg.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    const graphClient = new GraphClient(accessToken);
    await graphClient.moveMessage(msgId, "junkemail");
    Office.context.ui.closeContainer();
  } catch (error) {
    console.error("Moving to junk failed:", error);
    updateUIState(
      UIState.ERROR,
      error instanceof Error ? error.message : "Failed to move email to junk"
    );
  }
}
