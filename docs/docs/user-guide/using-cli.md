# Using the CLI for Microsoft 365

Information in this section will help you understand how the CLI for Microsoft 365 works and how you can use it most effectively.

## Start the CLI

To use CLI for Microsoft 365, execute specific commands directly from the command line.

<script src="https://asciinema.org/a/158207.js" id="asciicast-158207" async></script>

!!! attention
    When using the CLI for Microsoft 365, each CLI command must be prepended with `microsoft365` or `m365` for short. Without this, your shell will not know how the particular command should be executed.

Using the CLI for Microsoft 365 directly from the command line is invaluable if you want to write scripts consisting of a number of CLI for Microsoft 365 and other commands combined together. Additionally, you keep the access to all system commands and other CLIs available on your computer.

## List available commands

To list commands available with the CLI for Microsoft 365 type `help` in the CLI prompt, or `m365 help` directly in your shell.

Commands in the CLI for Microsoft 365 are combined into groups. You can list the commands available in the particular group by typing `help <group>`, for example `help spo` to list all commands related to SharePoint Online, or `m365 help spo` directly in your shell.

<script src="https://asciinema.org/a/158209.js" id="asciicast-158209" async></script>

## Get command help

Each command in the CLI for Microsoft 365 comes with help describing the command's purpose, available options and any other relevant details as well as related resources. To get the basic help information with command's description and available options, type `help <command>` or `m365 help <command>` in the shell, for example to get help for the `spo cdn get` command, type in the shell `m365 help spo cdn get`.

To get the complete help information including background information, examples and links to related information, use the `--help` option, for example `m365 spo cdn get --help`. This ability is also useful if you've already typed some options and don't want to lose your input but want to access the help, for example: `m365 spo cdn get --type Private --help`.

<script src="https://asciinema.org/a/158212.js" id="asciicast-158212" async></script>

## Required and optional command options

Commands in the CLI for Microsoft 365 often contain options that determine what the particular command should do. Command options vary from the URL of the site for which you would like to retrieve more information, to the type of Microsoft 365 CDN that you would like to manage.

Some options are required and necessary for the particular command to execute, while other are optional. When listing available options for the particular command, CLI for Microsoft 365 follows the naming convention where required options are wrapped in angle brackets (`< >`) while optional options are wrapped in square brackets (`[ ]`). For example, in the `spo cdn origin add` command, the origin you want to add is required (`-r, --origin <origin>`), while the type of CDN for which this origin should be added is optional (`-t, --type [type]`) and its value defaults to `Public`.

## Values with quotes

In cases, when the option's value contains spaces, it should be wrapped in quotes. For example, to create a modern team site for the _CLI for Microsoft 365_ team, you would execute in the shell:

```sh
m365 spo site add --alias office365cli --title "CLI for Microsoft 365"
```

When the value, that you want to provide contains quotes, it needs to be wrapped in quotes as well, for example to pass a JSON value in the CLI prompt, you would execute:

```sh
spo sitescript add --title "Contoso" --description "Contoso theme script" --content '{"abc": "def"}'
```

If you use the CLI for Microsoft 365 in Bash, the outer pair of quotes will be processed by Bash so the value needs to be wrapped in an additional pair of quotes, for example:

```sh
m365 spo sitescript add --title "Contoso" --description "Contoso theme script" --content '`{"abc": "def"}`'
```

## Verbose and debug mode

By default, commands output only the information returned by the corresponding Microsoft 365 API, whether the command result or error. You can choose for a more user-friendly output by using the `--verbose` option or setting the `OFFICE365CLI_VERBOSE` environment variable to `1`. For example: by default, when checking status of the Microsoft 365 Public CDN, you would see:

```sh
$ m365 spo cdn get
true
```

After adding the `--verbose` option, the output would change to:

```sh
$ m365 spo cdn get --verbose
Retrieving status of Public CDN...
Public CDN at https://contoso-admin.sharepoint.com is enabled
```

If you're experiencing problems when using the CLI for Microsoft 365, you can use the `--debug` option or set the `OFFICE365CLI_DEBUG` environment variable to `1`. On top of the output from the verbose mode, the debug mode will provide you with detailed information about all requests and responses from the Microsoft 365 APIs used by the command.

## Command completion

To help you use its commands, CLI for Microsoft 365 offers you the ability to autocomplete commands and options that you're typing in the prompt. Some additional setup, specific for the shell and terminal that you use, is required to enable command completion for CLI for Microsoft 365. For more information on configuring command completion for the CLI for Microsoft 365 see the [command completion](../concepts/completion.md) article.
