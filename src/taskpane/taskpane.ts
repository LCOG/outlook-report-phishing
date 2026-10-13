/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the
 * MIT license. See LICENSE in the project root for license information.
 */

import { ensureOfficeReady } from "../utils/office-type-guards";

import type { User } from "@microsoft/microsoft-graph-types";

type AuthManager = {
  initialize(): Promise<void>;
  ssoGetAccessToken(scopes: string[]): Promise<string>;
};

type MailService = {
  getRestItemId(): Promise<string>;
  getTargetMailbox(): Promise<string | undefined>;
  displayDialog(url: string, options: { height: number; width: number }): Promise<any>;
};

async function initializeFluentUI(): Promise<void> {
  const {
    provideFluentDesignSystem,
    fluentButton,
    fluentDivider,
    fluentRadio,
    fluentRadioGroup,
    fluentTextArea,
    fluentTooltip,
    fluentProgressRing,
  } = await import("@fluentui/web-components");

  provideFluentDesignSystem().register(
    fluentButton(),
    fluentDivider(),
    fluentRadio(),
    fluentRadioGroup(),
    fluentTextArea(),
    fluentTooltip(),
    fluentProgressRing()
  );
}

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
  reportMessageTypeGroup: (HTMLElement & { value: string }) | null;
  additionalInfoArea: (HTMLElement & { value: string }) | null;
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
let accountManager: AuthManager | null = null;
let officeMailService: MailService | null = null;

async function getAccountManager(): Promise<AuthManager> {
  if (!accountManager) {
    const { AccountManager } = await import("./authConfig");
    accountManager = new AccountManager();
  }

  return accountManager;
}

async function getOfficeMailService(): Promise<MailService> {
  if (!officeMailService) {
    const { OfficeMailService } = await import("../services/office-mail");
    officeMailService = new OfficeMailService();
  }

  return officeMailService;
}

function updateActionButtons(): void {
  const isSpamSelected = elements.reportMessageTypeGroup?.value === "spam";
  elements.additionalInfoArea?.classList.toggle("hidden", isSpamSelected);
  elements.reportEmailButton?.classList.toggle("hidden", isSpamSelected);
  elements.moveToJunkButton?.classList.toggle("hidden", !isSpamSelected);
}

function initializeElements(): void {
  // Cache all DOM element references with proper type assertions
  elements.reportEmailButton = document.getElementById("reportEmailButton");
  elements.reportEmailTooltip = document.getElementById("reportEmailTooltip");
  elements.reportInProgress = document.getElementById("reportInProgress");
  elements.reportSuccess = document.getElementById("reportingSuccess");
  elements.reportError = document.getElementById("reportingError");
  elements.reportErrorMessage = document.getElementById(
    "reportingErrorMessage"
  );
  elements.moveToJunkButton = document.getElementById("moveToJunkButton");
  elements.reportMessageTypeGroup = document.getElementById(
    "reportMessageTypeGroup"
  ) as (HTMLElement & { value: string }) | null;
  elements.additionalInfoArea = document.getElementById("additionalInfo") as
    | (HTMLElement & { value: string })
    | null;
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
  const results = [
    elements.reportInProgress, elements.reportSuccess, elements.reportError
  ];
  results.forEach((el) => el?.classList.add("hidden"));

  elements.reportEmailButton?.removeAttribute("disabled");

  // Show the appropriate element based on the current state
  switch (state) {
    case UIState.REPORTING:
      elements.reportEmailButton?.setAttribute("disabled", "true");
      if (elements.reportEmailTooltip) {
        elements.reportEmailTooltip.innerText =
          "Report in progress, please wait...";
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
      if (elements.reportEmailTooltip) {
        elements.reportEmailTooltip.innerText =
          "Click to retry reporting the email";
      }
      elements.reportError?.classList.remove("hidden");
      if (elements.reportErrorMessage) {
        elements.reportErrorMessage.innerText =
          errorMessage || "Something went wrong";
      }
      break;
    case UIState.IDLE:
      // Everything hidden is the default idle state
      if (elements.reportEmailTooltip) {
        elements.reportEmailTooltip.innerText =
          "Report this email as a phishing or a spam message";
      }
      break;
  }

  if (state !== UIState.SUCCESS) {
    updateActionButtons();
  }
}

function handleReportMessageTypeChange(): void {
  updateActionButtons();
}

// Initialize when Office is ready.
void (async () => {
  try {
    await ensureOfficeReady();
    await initializeFluentUI();
    initializeElements();

    if (elements.sideloadMsg) elements.sideloadMsg.style.display = "none";
    if (elements.appBody) elements.appBody.style.display = "flex";

    if (elements.reportEmailButton) {
      elements.reportEmailButton.addEventListener("click", handleReportClick);
    }
    if (elements.moveToJunkButton) {
      elements.moveToJunkButton.addEventListener(
        "click", handleMoveToJunkClick
      );
    }
    if (elements.reportMessageTypeGroup) {
      elements.reportMessageTypeGroup.addEventListener(
        "change", handleReportMessageTypeChange
      );
    }

    // Initialize MSAL.
    await (await getAccountManager()).initialize();
    updateUIState(UIState.IDLE);
  } catch (error) {
    console.error("Office initialization failed:", error);
  }
})();

/**
 * Gets the user's name and email
 */
async function getUserData(): Promise<User> {
  const accessToken = await (await getAccountManager()).ssoGetAccessToken(["user.read"]);
  const { GraphClient } = await import("../services/graph-client");
  const graphClient = new GraphClient(accessToken);
  return graphClient.getUser();
}

async function handleReportClick(): Promise<void> {
  try {
    updateUIState(UIState.REPORTING);

    const manager = await getAccountManager();
    const accessToken: string = await manager.ssoGetAccessToken(["mail.read", "mail.send"]);
    const mailService = await getOfficeMailService();
    const messageId = await mailService.getRestItemId();
    const targetMailbox = await mailService.getTargetMailbox();
    const user = await getUserData();

    const { GraphClient } = await import("../services/graph-client");
    const { PhishingReportService } = await import("../services/phishing-report");
    const graphClient = new GraphClient(accessToken, targetMailbox);
    const phishingService = new PhishingReportService(
      graphClient, reportPhishApiUrl as string
    );

    // Get report details from UI
    const reportType = elements.reportMessageTypeGroup?.value ?? "unknown";
    const additionalInfo = elements.additionalInfoArea?.value ?? "";

    // Forward email and log report in Team App
    const result = await phishingService.reportPhishing({
      messageId,
      user,
      reportType,
      additionalInfo,
      forwardToEmail: forwardEmail as string,
    });

    // Update UI based on result
    if (result.success) {
      updateUIState(UIState.SUCCESS);
    } else {
      updateUIState(UIState.ERROR, result.error ?? "Failed to report email");
    }
  } catch (error) {
    console.error("Report failed:", error);
    updateUIState(
      UIState.ERROR,
      error instanceof Error ? error.message : "Reporting the email failed"
    );
  }
}

async function handleMoveToJunkClick(): Promise<void> {
  try {
    const manager = await getAccountManager();
    const accessToken: string = await manager.ssoGetAccessToken([
      "mail.readwrite", "mail.send"
    ]);
    const mailService = await getOfficeMailService();
    const messageId = await mailService.getRestItemId();
    const targetMailbox = await mailService.getTargetMailbox();

    const { GraphClient } = await import("../services/graph-client");
    const { PhishingReportService } = await import("../services/phishing-report");
    const graphClient = new GraphClient(accessToken, targetMailbox);
    const phishingService = new PhishingReportService(
      graphClient, reportPhishApiUrl as string
    );

    await phishingService.moveToJunk(messageId);
    Office.context.ui.closeContainer();
  } catch (error) {
    console.error("Moving to junk failed:", error);
    updateUIState(
      UIState.ERROR,
      error instanceof Error ? error.message : "Failed to move email to junk"
    );
  }
}
