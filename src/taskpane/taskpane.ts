/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console */

import { AccountManager } from "./authConfig";
import { makeGraphRequest, makePostGraphRequest } from "./msgraph-helper";

const accountManager = new AccountManager();
const sideloadMsg = document.getElementById("sideload-msg");
const appBody = document.getElementById("app-body");
const getUserDataButton = document.getElementById("getUserData");
const reportEmailButton = document.getElementById("reportPhishingEmail");
const userName = document.getElementById("userName");
const userEmail = document.getElementById("userEmail");
const reportingSuccess = document.getElementById("reportingSuccess");
const reportingError = document.getElementById("reportingError");
const reportingErrorMessage = document.getElementById("reportingErrorMessage");

// Initialize when Office is ready.
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    if (getUserDataButton) {
      getUserDataButton.addEventListener("click", getUserData);
    }
    if (reportEmailButton) {
      reportEmailButton.addEventListener("click", reportPhishingEmail);
    }
    // Initialize MSAL.
    accountManager.initialize();
  }
});

/**
 * Gets the user data such as name and email and displays it
 * in the task pane.
 */
async function getUserData() {
  const userDataElement = document.getElementById("userData");
  // Specify minimum scopes for the token needed.
  const accessToken = await accountManager.ssoGetAccessToken(["user.read"]);

  const response: { displayName: string; mail: string } = await makeGraphRequest({
    accessToken,
    path: "/me",
  });

  if (userDataElement) {
    userDataElement.style.visibility = "visible";
  }
  if (userName) {
    userName.innerText = response.displayName ?? "";
  }
  if (userEmail) {
    userEmail.innerText = response.mail ?? "";
  }

  return response;
}

function displaySuccess() {
  if (reportingSuccess) {
    reportingSuccess.style.visibility = "visible";
  }
}

function displayError(errorMessage) {
  if (reportingError) {
    reportingError.style.visibility = "visible";
  }
  if (reportingErrorMessage) {
    reportingErrorMessage.innerText = errorMessage ?? "Something went wrong";
  }
}

async function logPhishingReport(emailAddress, message) {
  console.log("Got message:");
  console.log(message.body);

  try {
    const response = await fetch("http://localhost:8000/api/v1/phishreport", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        employee_email: emailAddress,
        email_message: message.body.content,
      }),
    });

    const result = await response.json();
    console.log(result);
  } catch (error) {
    console.error(error.message);
  }
}

async function reportPhishingEmail() {
  // Get access token for forwarding the email via Graph API
  const accessToken: string = await accountManager.ssoGetAccessToken(["mail.read", "mail.send"]);

  // The type of Office.context.mailbox.item can vary depending on the context
  // so we explicitly assign the expected type here.
  // TODO: Note that this is a suspected phishing email, and any links and
  // attachments on it need to be treated as malicious.
  // Research best practices on handling such emails.
  const msg: Office.MessageRead | undefined = Office.context.mailbox.item;

  console.log(msg);
  // TODO: display failure to retrieve the current email in the UI
  if (typeof msg === undefined) {
    console.error("Failed to retrieve current email from Office context.");
    return;
  }

  const { displayName, mail } = await getUserData();

  // Get the current email subject and body via Graph API and send it to
  // Team App phishing report endpoint.
  await makeGraphRequest({
    accessToken: accessToken,
    path: `/me/messages/${msg?.itemId}`,
    queryParams: "?$select=subject,body",
    additionalHeaders: { Prefer: 'outlook.body-content-type="text"' },
  })
    .then((currentMessageBody) => logPhishingReport(mail, currentMessageBody))
    .catch((reject) => console.error(reject));

  // Set body for the forwarding request
  const forwardBody = {
    comment: `${displayName} forwarded a suspicious email via the Report Phish add-in:`,
    toRecipients: [
      {
        emailAddress: {
          name: "LCOG IT",
          address: "ssavolainen@lcog-or.gov",
        },
      },
    ],
  };

  // Forward the current email to specified LCOG IT inbox via Graph API
  const forwardResult = await makePostGraphRequest({
    accessToken,
    path: `/me/messages/${msg?.itemId}/forward`,
    additionalHeaders: { "Content-Type": "application/json" },
    body: JSON.stringify(forwardBody),
  });

  if (forwardResult.ok) {
    displaySuccess();
  } else {
    console.error(forwardResult);
    displayError(`HTTP ${forwardResult.status} ${forwardResult.statusText}`);
  }

  // TODO: move email to junk folder after forwarding
  // Is any other cleanup needed?
  await makePostGraphRequest({
    accessToken,
    path: `/me/messages/${msg?.itemId}/move`,
    additionalHeaders: { "Content-Type": "application/json" },
    body: JSON.stringify({ destinationId: "junkemail" }),
  });
}
