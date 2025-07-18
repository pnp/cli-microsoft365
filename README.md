<h1 align="center">
  <a href="https://pnp.github.io/cli-microsoft365">
    <img alt="CLI for Microsoft 365" src="./docs/docs/images/pnp-cli-microsoft365-blue.svg" height="78">
  </a>
  <br>CLI for Microsoft 365<br>
</h1>

<h4 align="center">
  One CLI for Microsoft 365
</h4>

<p align="center">
  <a href="https://aka.ms/cli-m365/discord">
    <img src="https://img.shields.io/badge/Discord-aka.ms/cli--m365/discord-7289da?style=flat-square"
      alt="Discord" />
  </a>

  <a href="https://bsky.app/profile/climicrosoft365.bsky.social">
    <img src="https://img.shields.io/badge/Bsky-%40climicrosoft365.bsky.social-208bfe?style=flat-square"
      alt="Bluesky" />
  </a>

  <a href="https://x.com/climicrosoft365">
    <img src="https://img.shields.io/badge/X-%40climicrosoft365-blue?style=flat-square"
      alt="X" />
  </a>
</p>

<p align="center">
  <a href="https://github.com/pnp/cli-microsoft365/actions?query=workflow%3A%22Release+next%22">
    <img src="https://github.com/pnp/cli-microsoft365/workflows/Release%20next/badge.svg"
      alt="GitHub" />
  </a>
  
  <a href="https://www.npmjs.com/package/@pnp/cli-microsoft365">
    <img src="https://img.shields.io/npm/v/@pnp/cli-microsoft365/latest?style=flat-square"
      alt="npm @pnp/cli-microsoft365@latest" />
  </a>

  <a href="https://www.npmjs.com/package/@pnp/cli-microsoft365">
    <img src="https://img.shields.io/npm/v/@pnp/cli-microsoft365/next?style=flat-square"
      alt="npm @pnp/cli-microsoft365@next" />
  </a>
</p>

<p align="center">CLI for Microsoft 365 helps you manage your Microsoft 365 tenant and SharePoint Framework projects.</p>

<p align="center">
  <a href="https://pnp.github.io/cli-microsoft365">Website</a> |
  <a href="#features">Features</a> |
  <a href="#install">Install</a> |
  <a href="#usage">Usage</a> |
  <a href="#build">Build</a> |
  <a href="#contribute">Contribute</a>
</p>
<p align="center">
  <a href="https://github.com/pnp/cli-microsoft365/blob/main/CODE_OF_CONDUCT.md">Code of Conduct</a> | 
  <a href="#need-help">Need help?</a> |
  <a href="#disclaimer">Disclaimer</a>
</p>
<p align="center">
  <a href="#microsoft-365--power-platform-community">Microsoft 365 & Power Platform Community</a>
</p>
<p align="center">
  <img alt="CLI for Microsoft 365" src="./docs/docs/images/cli-microsoft365.gif" style="max-height: 500px;max-width: 100%;height: auto;width: auto;object-fit: contain;" />
</p>

## Features

- Run on any OS
  - Linux
  - MacOS
  - Windows
- Run on any shell
  - Azure Cloud Shell
  - bash
  - cmder
  - PowerShell
  - zsh
- Unified login
  - Access all your Microsoft 365 workloads
- Supported workloads
  - Bookings
  - Microsoft Entra ID
  - Microsoft Teams
  - Microsoft To Do
  - Microsoft Viva
  - OneDrive
  - OneNote
  - Outlook
  - Planner
  - Power Automate
  - Power Apps
  - Power Platform
  - Purview
  - SharePoint Embedded
  - SharePoint Online
  - SharePoint Premium
  - To Do
- Supported authentication methods
  - Azure Managed Identity
  - Certificate
  - Client Secret
  - Device Code
  - Federated identity
  - Username and Password
- Manage your SharePoint Framework projects
  - Upgrade your projects
  - Check your environment compatibility

> Follow our [Bluesky](https://bsky.app/profile/climicrosoft365.bsky.social), or [X](https://x.com/climicrosoft365) account to keep yourself updated about new features, improvements, and bug fixes.

## Install

To use the CLI for Microsoft 365 you need [`Node.js`](https://nodejs.org). The CLI has been tested with Node.js versions 20 and higher, but we recommend you to use the Node.js LTS version available at the moment.

```
npm install -g @pnp/cli-microsoft365
```

<details>
  <summary>Install beta version β</summary>

  ```
  npm install -g @pnp/cli-microsoft365@next
  ```
</details>

<details>
  <summary>Alternate package managers 🧶</summary>

  ### yarn

  ```
  yarn global add @pnp/cli-microsoft365
  ```

  ### npx

  ```
  npx @pnp/cli-microsoft365
  ```
</details>

<details>
  <summary>Run CLI for Microsoft 365 in a Docker container 🐳</summary>

  ```
  docker run --rm -it m365pnp/cli-microsoft365:latest
  ```

  Checkout our [guide](https://pnp.github.io/cli-microsoft365/user-guide/run-cli-in-docker-container/) to learn more about how to run CLI for Microsoft 365 using Docker
</details>

## Usage

>Before logging in, you should create a custom Microsoft Entra application registration. Use the `m365 setup` command or refer to the [Using your own Microsoft Entra identity](https://pnp.github.io/cli-microsoft365/user-guide/using-own-identity/) guide.

Use the `login` command to start the Device Code login flow to authenticate with your Microsoft 365 tenant.

```sh
m365 login
```

>For alternative authentication methods and usage, refer to the [login](https://pnp.github.io/cli-microsoft365/cmd/login/) command documentation

List all commands using the global `--help` option.

```sh
m365 --help
```

Get command information and example usage using the global `--help` option.

```sh
m365 spo site get --help
```

Execute a command and output response as JSON.

```sh
m365 spo site get --url https://contoso.sharepoint.com
```

Filter responses and return custom objects using [JMESPath](https://jmespath.org/) queries using the global `--query`  option.

```sh
m365 spo site list --query '[?Template==`GROUP#0`].{Title:Title, Url:Url}'
```

Execute a command and output response as text using the global `--output` option.

```sh
m365 spo site get --url https://contoso.sharepoint.com --output text
```

> For more examples and usage, refer to the [command](https://pnp.github.io/cli-microsoft365/cmd/login/) and  [sample scripts](https://pnp.github.io/cli-microsoft365/sample-scripts/introduction/) documentation.

## Build

To build and run this CLI locally, you will need [`node`](https://nodejs.org) `>= 22.0.0` installed.

```sh
# Clone this repository
$ git clone https://github.com/pnp/cli-microsoft365

# Go into the repository
$ cd cli-microsoft365

# Install dependencies
$ npm install

# Build the CLI
$ npm run build

# Symlink your local CLI build
$ npm link
```

When you execute any `m365` command from the terminal, it will now use your local clone of the CLI.

## Contribute

We love to accept contributions.

If you want to get involved with helping us grow the CLI, whether that is suggesting or adding a new command, extending an existing command, or updating our documentation, we would love to hear from you.

Check out our [Contributing Guide](https://pnp.github.io/cli-microsoft365/contribute/contributing-guide) for detailed information on how to contribute to this project.

## Need Help?

<h4 align="center">
  Join our community
</h4>
<p align="center">
  <a href="https://aka.ms/cli-m365/discord">
    <img alt="Discord" src="./docs/docs/images/discord-logo.png" width="100"/>
  </a>
</p>

## Microsoft 365 & Power Platform Community

CLI for Microsoft 365 is a [Microsoft 365 & Power Platform Community](https://pnp.github.io) (PnP) project. Microsoft 365 & Power Platform Community is a virtual team consisting of Microsoft employees and community members focused on helping the community make the best use of Microsoft products. CLI for Microsoft 365 is an open-source project not affiliated with Microsoft and not covered by Microsoft support. If you experience any issues using the CLI, please submit an issue in the [issues list](https://github.com/pnp/cli-microsoft365/issues).

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

![telemetry](https://telemetry.sharepointpnp.com/cli-microsoft365/readme)
