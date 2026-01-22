// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file handls MSAL auth for the fallback dialog page. */

import { createStandardPublicClientApplication } from "@azure/msal-browser";

import { getTokenRequest } from "../msalcommon";
import { defaultScopes, msalConfig } from "../msalconfig";
import { createLocalUrl } from "../util";

import type { AuthDialogResult } from "../authConfig";
import type { AuthenticationResult, IPublicClientApplication } from "@azure/msal-browser";

// read querystring parameter
function getQueryParameter(param: string): string | null {
  const params = new URLSearchParams(window.location.search);
  return params.get(param);
}

async function sendDialogMessage(message: string): Promise<void> {
  await Office.onReady();
  Office.context.ui.messageParent(message);
}
async function returnResult(
  publicClientApp: IPublicClientApplication,
  authResult: AuthenticationResult
): Promise<void> {
  publicClientApp.setActiveAccount(authResult.account);

  const authDialogResult: AuthDialogResult = {
    accessToken: authResult.accessToken,
  };

  await sendDialogMessage(JSON.stringify(authDialogResult));
}

export async function initializeMsal(): Promise<void> {
  // Use standard Public Client instead of nested because this is a fallback path when nested app authentication isn't available.
  const publicClientApp = await createStandardPublicClientApplication(msalConfig);
  try {
    if (getQueryParameter("logout") === "1") {
      await publicClientApp.logoutRedirect({
        postLogoutRedirectUri: createLocalUrl("dialog.html?close=1"),
      });
      return;
    } else if (getQueryParameter("close") === "1") {
      await sendDialogMessage("close");
      return;
    }
    const result = await publicClientApp.handleRedirectPromise();

    if (result) {
      return returnResult(publicClientApp, result);
    }
  } catch (ex: unknown) {
    const authDialogResult: AuthDialogResult = {
      error: ex instanceof Error ? ex.name : "UnknownError",
    };
    await sendDialogMessage(JSON.stringify(authDialogResult));
    return;
  }

  try {
    if (publicClientApp.getActiveAccount()) {
      const result = await publicClientApp.acquireTokenSilent(
        getTokenRequest(defaultScopes, false)
      );
      return returnResult(publicClientApp, result);
    }
  } catch {
    /* empty */
  }

  void publicClientApp.acquireTokenRedirect(
    getTokenRequest(defaultScopes, true, createLocalUrl("dialog.html"))
  );
}

void initializeMsal();
