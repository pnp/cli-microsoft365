![SharePoint Patterns and Practices](https://devofficecdn.azureedge.net/media/Default/PnP/sppnp.png)

# Office 365 CLI

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

> Depending on which settings you want to manage you might need to connect either to your tenant admin site (URL with `-admin`) in it, or to a regular SharePoint site. For more information refer to the help of the command you want to use.

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