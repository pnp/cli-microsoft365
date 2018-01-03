# Office 365 CLI

The Office 365 CLI allows you to manage different settings of your Microsoft Office 365 tenant on any platform.

<script type="text/javascript" src="https://asciinema.org/a/TJORGWjhqrbOSOQHe7fh3c11S.js" id="asciicast-TJORGWjhqrbOSOQHe7fh3c11S" async></script>

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

## SharePoint Patterns and Practices

Office 365 CLI is an open-source project driven by the [SharePoint Patterns and Practices](https://aka.ms/sppnp) initiative. The project is built and managed publicly on GitHub at [https://github.com/SharePoint/office365-cli](https://github.com/SharePoint/office365-cli) and accepts community contributions. We would encourage you to try it and [tell us what you think](https://github.com/SharePoint/office365-cli/issues). We would also love your help! We have a number of feature requests that are a [good starting point](https://github.com/SharePoint/office365-cli/issues?q=is%3Aissue+is%3Aopen+label%3A%22good+first+issue%22) to contribute to the project.

_“Sharing is caring”_

SharePoint PnP team