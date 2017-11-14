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