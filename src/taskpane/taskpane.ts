/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console */

import { AccountManager } from "./authConfig";
import { makeGraphRequest } from "./msgraph-helper";

const accountManager = new AccountManager();
const sideloadMsg = document.getElementById("sideload-msg");
const appBody = document.getElementById("app-body");
const getUserDataButton = document.getElementById("getUserData");
const reportEmailButton = document.getElementById("reportPhishingEmail");
const userName = document.getElementById("userName");
const userEmail = document.getElementById("userEmail");

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

  const response: { displayName: string; mail: string } = await makeGraphRequest(
    accessToken,
    "/me",
    ""
  );

  if (userDataElement) {
    userDataElement.style.visibility = "visible";
  }
  if (userName) {
    userName.innerText = response.displayName ?? "";
  }
  if (userEmail) {
    userEmail.innerText = response.mail ?? "";
  }
}

async function reportPhishingEmail() {
  // const apiUrl = "http://localhost:8000/api/v1/phishreport";
  const apiUrl = "https://api.team.lcog.org/api/v1/phishreport";
  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        employee_email: "dwilson@lcog.org",
        email_message: '{subject: "Test from development add-in", body: "This is a fake email"}',
      }),
    });

    const result = await response.json();
    console.log(result);
  } catch (error) {
    console.error(error.message);
  }
}
