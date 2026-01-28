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

Start by cloning the repository and installing dependencies:

```sh
# Clone the repository
git clone git@github.com:LCOG/outlook-report-phishing.git
cd outlook-report-phishing
# Enable pnpm and install dependencies
corepack enable pnpm
pnpm install
```

Following the patterns in `.env.example`, create a file called
`.env.development`. Then run the development server:

```sh
# Start the development server
pnpm run dev-server
```

`webpack` builds the development bundle and starts a development server at
`https://localhost:3000`.

- Open Outlook and select a message. You should be able to open the add-in from
  the apps menu in the ribbon.

- For testing the reporting integration to Team App, run the Team App backend
  locally.

Since Outlook add-ins are required to communicate over HTTPS even when run on
localhost, webpack creates self-signed certificates when first running the
development server. You should also see a prompt to allow installing them in the
OS trust store.

## Deploying to production

Following the patterns in `.env.example`, create a file called `.env.production`
and run `pnpm build`. This builds the production bundle in the `/dist`
directory.

To deploy the bundle to production, upload it to the `reportphish.lcog-or.gov`
S3 bucket.

## Add-in manifest files

The project has two manifest files: `manifest-dev.xml` and `manifest-prod.xml`.
These are used for two separate custom app deployments in Microsoft 365 admin
center. The development version is deployed to a limited set of users, and the
production version to the entire organization.

Based on `NODE_ENV`, webpack copies the corresponding manifest file to the
output bundle.

When making changes to the manifest files, increment the version number and
upload the updated manifest in the admin center.
