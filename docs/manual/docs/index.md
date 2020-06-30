# Office 365 CLI

Using the Office 365 CLI, you can manage your Microsoft Office 365 tenant and SharePoint Framework projects on any platform. No matter if you are on Windows, macOS or Linux, using Bash, Cmder or PowerShell, using the Office 365 CLI you can configure Office 365, manage SharePoint Framework projects and build automation scripts.

<script src="https://asciinema.org/a/265151.js" id="asciicast-265151" async></script>

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

Start managing the settings of your Office 365 tenant by logging in to it, using the `login` command, for example:

```sh
o365 login
```

To list all available commands, type in the Office 365 CLI prompt `help`:

```sh
o365 help
```

See the [User Guide](user-guide/installing-cli.md) to learn more about the Office 365 CLI and its capabilities.

## SharePoint Patterns and Practices

Office 365 CLI is an open-source project driven by the [SharePoint Patterns and Practices](https://aka.ms/sppnp) initiative. The project is built and managed publicly on GitHub at [https://github.com/pnp/office365-cli](https://github.com/pnp/office365-cli) and accepts community contributions. We would encourage you to try it and [tell us what you think](https://github.com/pnp/office365-cli/issues). We would also love your help! We have a number of feature requests that are a [good starting point](https://github.com/pnp/office365-cli/issues?q=is%3Aissue+is%3Aopen+label%3A%22good+first+issue%22) to contribute to the project.

_“Sharing is caring”_

SharePoint PnP team