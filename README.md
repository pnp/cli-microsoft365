![SharePoint Patterns and Practices](https://devofficecdn.azureedge.net/media/Default/PnP/sppnp.png)

@latest (master) | @next (dev)
:--------------: | :---------:
[![CircleCI](https://circleci.com/gh/SharePoint/office365-cli/tree/master.svg?style=shield&circle-token=ce99e8046a231e1959248a61e7e32f9ae1abc8cf)](https://circleci.com/gh/SharePoint/office365-cli/tree/master)|[![CircleCI](https://circleci.com/gh/SharePoint/office365-cli/tree/dev.svg?style=shield&circle-token=ce99e8046a231e1959248a61e7e32f9ae1abc8cf)](https://circleci.com/gh/SharePoint/office365-cli/tree/dev)
[![Coverage Status](https://coveralls.io/repos/github/SharePoint/office365-cli/badge.svg?branch=master)](https://coveralls.io/github/SharePoint/office365-cli?branch=master)|[![Coverage Status](https://coveralls.io/repos/github/SharePoint/office365-cli/badge.svg?branch=dev)](https://coveralls.io/github/SharePoint/office365-cli?branch=dev)
[![bitHound Dependencies](https://www.bithound.io/github/SharePoint/office365-cli/badges/dependencies.svg)](https://www.bithound.io/github/SharePoint/office365-cli/master/dependencies/npm)|[![bitHound Dependencies](https://www.bithound.io/github/SharePoint/office365-cli/badges/dependencies.svg)](https://www.bithound.io/github/SharePoint/office365-cli/dev/dependencies/npm)

# Office 365 CLI

[![Join the chat at https://gitter.im/office365-cli/cli](https://badges.gitter.im/Join%20Chat.svg)](https://gitter.im/office365-cli/cli?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

The Office 365 CLI allows you to manage different settings of your Microsoft Office 365 tenant on any platform.

[![asciicast](https://asciinema.org/a/TJORGWjhqrbOSOQHe7fh3c11S.png)](https://asciinema.org/a/TJORGWjhqrbOSOQHe7fh3c11S)

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

Start managing the settings of your Office 365 tenant by connecting to it, using the `spo connect <url>` site, for example:

```sh
o365$ spo connect https://contoso-admin.sharepoint.com
```

> Depending on which settings you want to manage you might need to connect either to your tenant admin site (URL with `-admin` in it), or to a regular SharePoint site. For more information refer to the help of the command you want to use.

To list all available commands, type in the Office 365 CLI prompt `help`:

```sh
o365$ help
```

To exit the CLI, type `exit`:

```sh
o365$ exit
```

## Sharing is Caring

We'd love your help! If you have ideas for new features or feedback, let us know by creating an issue in the [issues list](https://github.com/SharePoint/office365-cli/issues). Before you submit a PR with your improvements, please review our [project guides](./docs/guides/index.md).

## Legal

This project welcomes contributions and suggestions. Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

![telemetry](https://telemetry.sharepointpnp.com/office365-cli/readme)