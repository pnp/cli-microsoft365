# CLI for Microsoft 365

Using the CLI for Microsoft 365, you can manage your Microsoft 365 tenant and SharePoint Framework projects on any platform. No matter if you are on Windows, macOS or Linux, using Bash, Cmder or PowerShell, using the CLI for Microsoft 365 you can configure Microsoft 365, manage SharePoint Framework projects and build automation scripts.

<script id="asciicast-445653" src="https://asciinema.org/a/445653.js" async></script>

## Installation

The CLI for Microsoft 365 is distributed as an NPM package. To use it, install it globally using:

```sh
npm i -g @pnp/cli-microsoft365
```

or using yarn:

```sh
yarn global add @pnp/cli-microsoft365
```

## Getting started

Start managing the settings of your Microsoft 365 tenant by logging in to it, using the `login` command, for example:

```sh
m365 login
```

To list all available commands, type in the CLI for Microsoft 365 prompt `help`:

```sh
m365 help
```

See the [User Guide](user-guide/installing-cli.md) to learn more about the CLI for Microsoft 365 and its capabilities.

## Microsoft 365 Platform Community

CLI for Microsoft 365 is a [Microsoft 365 Platform Community](https://pnp.github.io) (PnP) project. Microsoft 365 Platform Community is a virtual team consisting of Microsoft employees and community members focused on helping the community make the best use of Microsoft products. CLI for Microsoft 365 is an open-source project not affiliated with Microsoft and not covered by Microsoft support. If you experience any issues using the CLI, please submit an issue in the [issues list](https://github.com/pnp/cli-microsoft365/issues).

The project is built and managed publicly on GitHub at [https://github.com/pnp/cli-microsoft365](https://github.com/pnp/cli-microsoft365) and accepts community contributions. We would encourage you to try it and [tell us what you think](https://github.com/pnp/cli-microsoft365/issues). We would also love your help! We have a number of feature requests that are a [good starting point](https://github.com/pnp/cli-microsoft365/issues?q=is%3Aissue+is%3Aopen+label%3A%22good+first+issue%22) to contribute to the project.
