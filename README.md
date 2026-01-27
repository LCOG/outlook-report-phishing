# Outlook add-in for reporting phishing emails

![Icon with a cartoon anglerfish holding a mail envelope](./assets/rp-icon-retouch-256.png)

Report phishing emails to IT and log the reports in [Team App](https://github.com/LCOG/lcog-team).

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

## Developing the add-in

To develop the add-in, run the development server.

```sh
# Clone the repository
git clone git@github.com:LCOG/outlook-report-phishing.git
cd outlook-report-phishing
# Enable pnpm and install dependencies
corepack enable pnpm
pnpm install
# Start the development server
pnpm run dev-server
```

Open Outlook and select a message. You should be able to open the add-in from
the apps menu in the ribbon.

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

For more information on debugging with VS Code, see
[Debugging](https://code.visualstudio.com/Docs/editor/debugging). For more
information on debugging Office Add-ins in VS Code, see [Debug Office Add-ins
on Windows using Visual Studio Code and Microsoft Edge WebView2
(Chromium-based)](https://learn.microsoft.com/office/dev/add-ins/testing/debug-desktop-using-edge-chromium)
