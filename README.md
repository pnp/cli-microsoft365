<img src="./docs/manual/docs/images/pnp-office365-cli-blue.svg" alt="Office 365 CLI" height=78 />

@latest (master) | @next (dev)
:--------------: | :---------:
[![CircleCI](https://circleci.com/gh/pnp/office365-cli/tree/master.svg?style=shield&circle-token=ce99e8046a231e1959248a61e7e32f9ae1abc8cf)](https://circleci.com/gh/pnp/office365-cli/tree/master)|[![CircleCI](https://circleci.com/gh/pnp/office365-cli/tree/dev.svg?style=shield&circle-token=ce99e8046a231e1959248a61e7e32f9ae1abc8cf)](https://circleci.com/gh/pnp/office365-cli/tree/dev)
[![Coverage Status](https://coveralls.io/repos/github/pnp/office365-cli/badge.svg?branch=master)](https://coveralls.io/github/pnp/office365-cli?branch=master)|[![Coverage Status](https://coveralls.io/repos/github/pnp/office365-cli/badge.svg?branch=dev)](https://coveralls.io/github/pnp/office365-cli?branch=dev)

# Office 365 CLI

Using the Office 365 CLI, you can manage your Microsoft Office 365 tenant and SharePoint Framework projects on any platform. No matter if you are on Windows, macOS or Linux, using Bash, Cmder or PowerShell, using the Office 365 CLI you can configure Office 365, manage SharePoint Framework projects and build automation scripts.

[![asciicast](https://asciinema.org/a/265151.png)](https://asciinema.org/a/265151)

## Installation

The Office 365 CLI is distributed as an NPM package. To use it, install it globally using:

```sh
npm i -g @pnp/office365-cli
```

or using yarn:

```sh
yarn global add @pnp/office365-cli
```

The beta version of the Office 365 CLI can be installed by using the `@next` tag:

```sh
npm i -g @pnp/office365-cli@next
```

## Getting started

Start the Office 365 CLI by typing in the command line:

```sh
$ office365

o365$ _
```

Running the `office365` command will start the immersive CLI with its own command prompt.

Start managing the settings of your Office 365 tenant by logging in to it, using the `login` command, for example:

```sh
o365$ login
```

> Office 365 CLI will automatically detect the URL of your tenant based on the account that you use to sign in.

To list all available commands, type in the Office 365 CLI prompt `help`:

```sh
o365$ help
```

To exit the CLI, type `exit`:

```sh
o365$ exit
```

See the [User Guide](docs/manual/docs/user-guide/installing-cli.md) to learn more about the Office 365 CLI and its capabilities.

## Sharing is Caring

We'd love your help! If you have ideas for new features or feedback, let us know by creating an issue in the [issues list](https://github.com/pnp/office365-cli/issues). Before you submit a PR with your improvements, please review our [project guides](./docs/guides/index.md).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

![SharePoint Patterns and Practices](https://devofficecdn.azureedge.net/media/Default/PnP/sppnp.png)

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

![telemetry](https://telemetry.sharepointpnp.com/office365-cli/readme)