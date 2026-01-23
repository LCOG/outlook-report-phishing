// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/* This file provides MSAL auth configuration to get access token through nested app authentication. */

import {
  BrowserAuthError,
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  type AuthenticationResult,
} from "@azure/msal-browser";

import { getTokenRequest } from "./msalcommon";
import { msalConfig } from "./msalconfig";
import { createLocalUrl } from "./util";
import { OfficeMailService } from "../services/office-mail";
import { ensureOfficeReady } from "../utils/office-type-guards";

export type AuthDialogResult = {
  accessToken?: string;
  error?: string;
};

type DialogEventMessage = { message: string; origin: string | undefined };
type DialogEventError = { error: number };
type DialogEventArg = DialogEventMessage | DialogEventError;

// Encapsulate functions for getting user account and token information.
export class AccountManager {
  private pca: IPublicClientApplication | undefined = undefined;
  private _dialogApiResult: Promise<string> | null = null;
  private _usingFallbackDialog = false;
  private officeMailService = new OfficeMailService();

  private getSignOutButton(): HTMLElement | null {
    return document.getElementById("signOutButton");
  }

  private setSignOutButtonVisibility(isVisible: boolean): void {
    const signOutButton = this.getSignOutButton();
    if (signOutButton) {
      signOutButton.style.visibility = isVisible ? "visible" : "hidden";
    }
  }

  private isNestedAppAuthSupported(): boolean {
    return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  }

  // Initialize MSAL public client application.
  async initialize(): Promise<void> {
    // Make sure office.js is initialized
    await ensureOfficeReady();

    // If auth is not working, enable debug logging to help diagnose.
    this.pca = await createNestablePublicClientApplication(msalConfig);

    // If Office does not support Nested App Auth provide a sign-out button since the user selects account
    if (!this.isNestedAppAuthSupported() && this.pca.getActiveAccount() !== null) {
      this.setSignOutButtonVisibility(true);
    }
    void this.getSignOutButton()?.addEventListener("click", () => {
      void this.signOut();
    });
  }

  private async signOut(): Promise<void> {
    if (this._usingFallbackDialog) {
      await this.signOutWithDialogApi();
    } else {
      await this.pca?.logoutPopup();
    }

    this.setSignOutButtonVisibility(false);
  }

  /**
   *
   * @param scopes the minimum scopes needed.
   * @returns An access token.
   */
  async ssoGetAccessToken(scopes: string[]): Promise<string> {
    if (this._dialogApiResult !== null) {
      return this._dialogApiResult;
    }

    if (this.pca === undefined) {
      throw new Error("AccountManager is not initialized!");
    }

    try {
      console.warn("Trying to acquire token silently...");
      const authResult: AuthenticationResult = await this.pca.acquireTokenSilent(
        getTokenRequest(scopes, false)
      );
      console.warn("Acquired token silently.");
      return authResult.accessToken;
    } catch (error) {
      console.warn(`Unable to acquire token silently: ${error}`);
    }

    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.warn("Trying to acquire token interactively...");
      const selectAccount = this.pca.getActiveAccount() === null;
      const authResult: AuthenticationResult = await this.pca.acquireTokenPopup(
        getTokenRequest(scopes, selectAccount)
      );
      console.warn("Acquired token interactively.");
      if (selectAccount) {
        this.pca.setActiveAccount(authResult.account);
      }
      if (!this.isNestedAppAuthSupported()) {
        this.setSignOutButtonVisibility(true);
      }
      return authResult.accessToken;
    } catch (popupError) {
      // Optional fallback if about:blank popup should not be shown
      if (popupError instanceof BrowserAuthError && popupError.errorCode === "popup_window_error") {
        const accessToken = await this.getTokenWithDialogApi();
        return accessToken;
      } else {
        // Acquire token interactive failure.
        console.error(`Unable to acquire token interactively: ${popupError}`);
        throw new Error(`Unable to acquire access token: ${popupError}`);
      }
    }
  }

  /**
   * Gets an access token by using the Office dialog API to handle authentication. Used for fallback scenario.
   * @returns The access token.
   */
  async getTokenWithDialogApi(): Promise<string> {
    this._dialogApiResult = (async () => {
      const dialog = await this.officeMailService.displayDialog(createLocalUrl("dialog.html"), {
        height: 60,
        width: 30,
      });

      return new Promise<string>((resolve, reject) => {
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: DialogEventArg) => {
          const errorArg = arg as DialogEventError;
          if (errorArg.error === 12006) {
            this._dialogApiResult = null;
            reject(new Error("Dialog closed"));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: DialogEventArg) => {
          const messageArg = arg as DialogEventMessage;
          const parsedMessage = JSON.parse(messageArg.message);
          void dialog.close();

          if (parsedMessage.accessToken !== undefined && parsedMessage.accessToken !== null) {
            resolve(parsedMessage.accessToken as string);
            this.setSignOutButtonVisibility(true);
            this._usingFallbackDialog = true;
          } else {
            const errorMessage = parsedMessage.error as string | undefined;
            if (errorMessage !== undefined && errorMessage !== "") {
              reject(new Error(errorMessage));
            } else {
              reject(new Error("Access token not found in dialog response"));
            }
            this._dialogApiResult = null;
          }
        });
      });
    })();

    return this._dialogApiResult;
  }

  async signOutWithDialogApi(): Promise<void> {
    const dialog = await this.officeMailService.displayDialog(
      createLocalUrl("dialog.html?logout=1"),
      {
        height: 60,
        width: 30,
      }
    );

    return new Promise((resolve) => {
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        this.setSignOutButtonVisibility(false);
        this._dialogApiResult = null;
        resolve();
        void dialog.close();
      });
    });
  }
}
