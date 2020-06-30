# Installing the Office 365 CLI

Thank you for your interest in the Office 365 CLI. Following information will help you install the Office 365 CLI and keep it up to date.

## Prerequisites

To use the Office 365 CLI you need Node.js. The CLI has been tested with Node.js versions 6 and higher, but we recommend you to use the Node.js LTS version available at the moment. For more information on installing Node.js for your platform visit [https://nodejs.org](https://nodejs.org).

Office 365 CLI works on Windows, macOS and Linux and you can use it with any shell on these platforms.

## Install the Office 365 CLI

Office 365 CLI is distributed as a Node.js package and published on [npmjs.com](https://www.npmjs.com). You can install it using your Node package manager, such as npm or Yarn.

To install the Office 365 CLI, in the command line execute:

```sh
npm install -g @pnp/office365-cli
```

<script src="https://asciinema.org/a/158191.js" id="asciicast-158191" async></script>

If you're using Yarn, you can install the Office 365 CLI by executing:

```sh
yarn global add @pnp/office365-cli
```

!!! tip
    We are regularly publishing beta versions of the Office 365 CLI with latest features and fixes. If you want to use the beta version of the CLI, add `@next` to the package name when installing the CLI, for example `npm install -g @pnp/office365-cli@next`. Please note, that you can have installed either the beta version or the public version of the CLI but not both.

## Check the installed version

To check which version of the Office 365 CLI you have installed on your computer, execute in the command line:

```sh
o365 version
```

Alternatively, you can check the version of the Node.js package you have installed, by executing:

```sh
npm ls -g --depth=0
```

The version of the CLI is the same as the version of the Node.js package distributing the CLI, so by using either of the commands you can control which version of the CLI you have installed.

## Check if a new version is available

We are continuously evolving the Office 365 CLI and adding more features to it. You can download new versions of the CLI from npmjs.com. To check, if a new version of the Office 365 CLI is available, execute in the command line:

```sh
npm outdated -g
```

For each package that you have installed globally, npm will show the version you have currently installed as well as the latest version available on npm.

If you want to check if a new beta version of the Office 365 CLI is available, execute in the command line `npm view @pnp/office365-cli`. Next, compare the version listed as the `@next` tag with the version you have installed. Beta versions of the Office 365 CLI are tagged with source code commits so that it's easy for the team to debug it in case of any issues.

```sh hl_lines="5"
$ npm view @pnp/office365-cli

{ name: '@pnp/office365-cli',
  description: 'CLI for managing Office 365 configuration',
  'dist-tags': { next: '0.5.0-beta.fe510b6', latest: '0.4.0' },
  versions:
  [ '0.1.0-beta.b35346a',
    '0.1.0-beta.b7db425',
    '0.1.0-beta.b85510d',
    '0.1.1-beta.25b1725',
    ...
```

## Update the CLI

To update the CLI, execute in the command line:

```sh
npm install -g @pnp/office365-cli@latest
```

This will download and install the latest public version of the Office 365 CLI. If you want to update to the latest beta version of the CLI, replace `@latest` with `@next`.

!!! important
    New version of the Office 365 CLI often contains new commands. Don't forget to update command completion for your terminal to get suggestions for the latest commands added in the CLI. For more information see the article on [command completion](../concepts/completion.md).

## Uninstall the CLI

!!! attention
    Before uninstalling the CLI, log out from Office 365 using the `logout` command. Office 365 CLI persists connection information on your computer and if you don't log out before uninstalling the CLI, this information will be left on your computer and you will have to remove it manually. For more information see the article on [persisting connection information](../concepts/persisting-connection.md)

To uninstall the Office 365 CLI, execute in the command line:

```sh
npm uninstall -g @pnp/office365-cli
```

This command will remove all Office 365 CLI files from your computer.

If you have configured command completion for the Office 365 CLI in your terminal, remove the completion following instructions specific to your terminal, to avoid errors in your terminal.