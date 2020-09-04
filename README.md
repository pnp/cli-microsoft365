<img src="./docs/docs/images/pnp-cli-microsoft365-blue.svg" alt="CLI for Microsoft 365" height=78 />

@latest/@next (master) |
:--------------: |
[![CircleCI](https://circleci.com/gh/pnp/cli-microsoft365/tree/master.svg?style=shield&circle-token=ce99e8046a231e1959248a61e7e32f9ae1abc8cf)](https://circleci.com/gh/pnp/cli-microsoft365/tree/master)|
[![Coverage Status](https://coveralls.io/repos/github/pnp/cli-microsoft365/badge.svg?branch=master)](https://coveralls.io/github/pnp/cli-microsoft365?branch=master)|

# CLI for Microsoft 365

Using CLI for Microsoft 365, you can manage your Microsoft 365 tenant and SharePoint Framework projects on any platform. No matter if you are on Windows, macOS or Linux, using Bash, Cmder or PowerShell, using CLI for Microsoft 365 you can configure Microsoft 365, manage SharePoint Framework projects and build automation scripts.

[![asciicast](https://asciinema.org/a/346365.png)](https://asciinema.org/a/346365)

## Installation

CLI for Microsoft 365 is distributed as an NPM package. To use it, install it globally using:

```sh
npm i -g @pnp/cli-microsoft365
```

or using yarn:

```sh
yarn global add @pnp/cli-microsoft365
```

The beta version of CLI for Microsoft 365 can be installed by using the `@next` tag:

```sh
npm i -g @pnp/cli-microsoft365@next
```

## Getting started

Start managing the settings of your Microsoft 365 tenant by logging in to it, using the `login` command, for example:

```sh
m365 login
```

> CLI for Microsoft 365 will automatically detect the URL of your tenant based on the account that you use to sign in.

To list all available commands, type in CLI for Microsoft 365 prompt `help`:

```sh
m365 help
```

See the [User Guide](docs/docs/user-guide/installing-cli.md) to learn more about CLI for Microsoft 365 and its capabilities.

## Sharing is Caring

We'd love your help! If you have ideas for new features or feedback, let us know by creating an issue in the [issues list](https://github.com/pnp/cli-microsoft365/issues). Before you submit a PR with your improvements, please review our [project guides](https://github.com/pnp/cli-microsoft365/wiki).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

![Microsoft 365 Patterns and Practices](./docs/docs/images/parker-ms-logo-300.png)

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

![telemetry](https://telemetry.sharepointpnp.com/cli-microsoft365/readme)
