# Outlook add-in for reporting phishing emails

Report phishing emails to IT and track reports in [LCOG Team App](https://github.com/LCOG/lcog-team).

This repository is based on [Microsoft's sample code](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO-NAA)
for implementing an Outlook add-in with SSO using nested app authentication.

## Prerequisites

- [Node.js](https://nodejs.org/) 24 LTS.
- [npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm) version 8 or greater.
- [pnpm](https://pnpm.io/) for package management.
- Office connected to a Microsoft 365 subscription (including Office on the web).
- Your user has access to the **LCOG Report Phish** app in
  **Microsoft 365 admin center > Settings > Integrated apps**.
  - So far we have not been successful in sideloading the add-in to Outlook, so
    we're loading the add-in through a limited app deployment.

## Build and run the solution

1. Clone the repository.
1. Run `pnpm install`.
1. Start the development server: `pnpm run dev-server`.
1. Open Outlook and select a message. You should be able to open the add-in
   from the apps menu in the ribbon.

For testing the reporting integration to Team App, run the Team App backend
locally.

## Choose a manifest type

By default, the sample uses an add-in only manifest. However, you can switch
the project between the add-in only manifest and the unified manifest. For more
information about the differences between them, see [Office Add-ins
manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).

### To switch to the Unified manifest for Microsoft 365

Copy all files from the **manifest-configurations/unified** subfolder to the
sample's root folder, replacing any existing files that have the same names. We
recommend that you delete the **manifest.xml** file from the root folder, so
only files needed for the unified manifest are present in the root.

### To switch back to the Add-in only manifest

If you want to switch back to the add-in only manifest, copy the files in the
**manifest-configurations/add-in-only** subfolder to the sample's root folder.
We recommend that you delete the **manifest.json** file from the root folder.

## Debugging steps

You can debug the sample by opening the project in VS Code.

1. Select the **Run and Debug** icon in the **Activity Bar** on the side of VS
   Code. You can also use the keyboard shortcut
   <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>D</kbd>.
1. Select the launch configuration you want from the **Configuration dropdown**
   in the **Run and Debug** view. For example, **Outlook Desktop (Edge
   Chromium)**.
1. Start your debug session with **F5**, or **Run** > **Start Debugging**.

![The VS Code debug view.](./assets/vs-code-debug-view.png)

For more information on debugging with VS Code, see
[Debugging](https://code.visualstudio.com/Docs/editor/debugging). For more
information on debugging Office Add-ins in VS Code, see [Debug Office Add-ins
on Windows using Visual Studio Code and Microsoft Edge WebView2
(Chromium-based)](https://learn.microsoft.com/office/dev/add-ins/testing/debug-desktop-using-edge-chromium)

## Key parts of this sample

The `src/taskpane/authConfig.ts` file contains the MSAL code for configuring
and using NAA. It contains a class named AccountManager which manages getting
user account and token information.

- The `initialize` function is called from Office.onReady to configure and
  initialize MSAL to use NAA.
- The `ssoGetAccessToken` function gets an access token for the signed in user
  to call Microsoft Graph APIs.
- The `getTokenWithDialogApi` function uses the Office dialog API to support a
  fallback option if NAA fails.

The `src/taskpane/taskpane.ts` file contains code that runs when the user
chooses buttons in the task pane. It uses the AccountManager class to get
tokens or user information depending on which button is chosen.

The `src/taskpane/msgraph-helper.ts` file contains code to construct and make a
REST call to the Microsoft Graph API.

### Fallback code

The `fallback` folder contains files to fall back to an alternate
authentication method if NAA is unavailable and fails. When your code calls
`acquireTokenSilent`, and NAA is unavailable, an error is thrown. The next step
is the code calls `acquireTokenPopup`. MSAL then attempts to sign in the user
by opening a dialog box with `window.open` and `about:blank`. Some older
Outlook clients don't support the `about:blank` dialog box and cause the
`aquireTokenPopup` method to fail. You can catch this error and fall back to
using the Office dialog API to open the auth dialog instead.

- the `src/taskpane/authconfig.ts` file contains the following code to detect
  the error and fall back to using the Office dialog API.

```typescript
    // Optional fallback if about:blank popup should not be shown
    if (popupError instanceof BrowserAuthError && popupError.errorCode === "popup_window_error") {
        const accessToken = await this.getTokenWithDialogApi();
        return accessToken;
```

- The `src/taskpane/fallback/fallbackauthdialog.ts` file contains code to
  initialize MSAL and acquire an access token. It sends the access token back to
  the task pane.
