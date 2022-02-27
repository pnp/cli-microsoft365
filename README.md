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
  <a href="https://github.com/pnp/cli-microsoft365/actions?query=workflow%3A%22Release+next%22">
    <img src="https://github.com/pnp/cli-microsoft365/workflows/Release%20next/badge.svg"
      alt="GitHub" />
  </a>

  <a href="https://twitter.com/climicrosoft365">
    <img src="https://img.shields.io/badge/Twitter-%40climicrosoft365-blue?style=flat-square"
      alt="Twitter" />
  </a>
</p>

<p align="center">
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
  <a href="#sharing-is-caring">Sharing is Caring</a> |
  <a href="#code-of-conduct">Code of Conduct</a> | 
  <a href="#disclaimer">Disclaimer</a>
</p>

<p align="center">
  <img alt="CLI for Microsoft 365" src="./docs/docs/images/cli-microsoft365.gif" height="500" />
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
  - Azure Active Directory
  - Microsoft Teams
  - Microsoft To Do
  - OneDrive
  - Outlook
  - Planner
  - Power Automate
  - Power Apps
  - Skype for Business
  - SharePoint Online
  - Yammer
- Supported authentication methods
  - Azure Managed Identity
  - Certificate
  - Client Secret
  - Device Code
  - Username and Password
- Manage your SharePoint Framework projects
  - Uprade your projects
  - Check your environment compatibility

> Follow our [Twitter](https://twitter.com/climicrosoft365) account to keep yourself updated about new features, improvements, and bug fixes.

## Install

To install this CLI, you will need [`node`](https://nodejs.org) `>= 8.0.0` installed.

```
npm install -g @pnp/cli-microsoft365
```

<details>
  <summary>Install beta version  β</summary>

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

Use the `login` command to start the Device Code login flow to authenticate with your Microsoft 365 tenant. 

```sh
m365 login
```

>On your first login you will be asked to consent to several permissions that the `PnP Management Shell` multi-tenant app requires for the commands to work successfully against your tenant. If you want to create your own identity to use with the CLI, refer to the [Using your own Azure AD Identity](https://pnp.github.io/cli-microsoft365/user-guide/using-own-identity/) guide.

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

> For more examples and usage, refer to the [command](https://pnp.github.io/cli-microsoft365/cmd/login/) and  [sample scripts](https://pnp.github.io/cli-microsoft365/sample-scripts/) documentation.

## Build

To build and run this CLI locally, you will need [`node`](https://nodejs.org) `>= 16.0.0` installed.

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

If you want to get involved with helping us grow the CLI, whether that is suggesting or adding a new command, extending an existing command or updating our documentation, we would love to hear from you.

Checkout our [Wiki](https://github.com/pnp/cli-microsoft365/wiki) for guides on how to contribute to this project.

## Microsoft 365 Platform Community

CLI for Microsoft 365 is a [Microsoft 365 Platform Community](https://pnp.github.io) (PnP) project. Microsoft 365 Platform Community is a virtual team consisting of Microsoft employees and community members focused on helping the community make the best use of Microsoft products. CLI for Microsoft 365 is an open-source project not affiliated with Microsoft and not covered by Microsoft support. If you experience any issues using the CLI, please submit an issue in the [issues list](https://github.com/pnp/cli-microsoft365/issues).

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

![telemetry](https://telemetry.sharepointpnp.com/cli-microsoft365/readme)
