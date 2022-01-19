# Use the CLI for Microsoft 365

Information in this section will help you understand how the CLI for Microsoft 365 works and how you can use it most effectively.

## Start the CLI

To use CLI for Microsoft 365, execute specific commands directly from the command line.

<script id="asciicast-445654" src="https://asciinema.org/a/445654.js" async></script>

!!! attention
    When using the CLI for Microsoft 365, each CLI command must be prepended with `microsoft365` or `m365` for short. Without this, your shell will not know how the particular command should be executed.

Using the CLI for Microsoft 365 directly from the command line is invaluable if you want to write scripts consisting of a number of CLI for Microsoft 365 and other commands combined together. Additionally, you keep the access to all system commands and other CLIs available on your computer.

## List available commands

To list commands available with the CLI for Microsoft 365 type `help` in the CLI prompt, or `m365 help` directly in your shell.

Commands in the CLI for Microsoft 365 are combined into groups. You can list the commands available in the particular group by typing `help <group>`, for example `help spo` to list all commands related to SharePoint Online, or `m365 help spo` directly in your shell.

<script id="asciicast-445655" src="https://asciinema.org/a/445655.js" async></script>

## Get command help

Each command in the CLI for Microsoft 365 comes with help describing the command's purpose, available options and any other relevant details as well as related resources. To get the basic help information with command's description and available options, type `help <command>` or `m365 help <command>` in the shell, for example to get help for the `spo cdn get` command, type in the shell `m365 help spo cdn get`.

To get the complete help information including background information, examples and links to related information, use the `--help` option, for example `m365 spo cdn get --help`. This ability is also useful if you've already typed some options and don't want to lose your input but want to access the help, for example: `m365 spo cdn get --type Private --help`.

<script id="asciicast-445656" src="https://asciinema.org/a/445656.js" async></script>

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
m365 spo sitescript add --title "Contoso" --description "Contoso theme script" --content '{"abc": "def"}'
```

## Working with SharePoint URLs in `spo` commands

CLI for Microsoft 365 contains a number of commands for managing SharePoint Online. Each of these commands requires you to specify the site or web on which you want to execute the command. For example, to get information about a site collection located at `https://contoso.sharepoint.com/sites/contoso`, you'd execute:

```sh
m365 spo site get --url https://contoso.sharepoint.com/sites/contoso
```

If you executed an `spo` command previously, CLI for Microsoft 365 already knows the hostname of your SharePoint Online tenant. In such case, you can use a server-relative URL as well:

```sh
m365 spo site get --url /sites/contoso
```

If you try to use a server-relative URL but CLI for Microsoft 365 doesn't know of your SharePoint Online URL yet, you will see an error prompting you to either use an absolute URL or set the SPO URL using the `spo set` command:

```sh
m365 spo set --url https://contoso.sharepoint.com
```

You can also execute a command like `m365 spo site list` that will automatically detect your SharePoint Online tenant URL for you.

To check if CLI detected the SPO URL previously, use the `m365 spo get` command.

## Passing complex content into CLI options

When passing complex content into CLI options, such as JSON strings, you will need to properly escape nested quotes. The exact way to do it, depends on the shell that you're using. Alternatively, you can choose to pass complex content by storing the complex content in a file and passing the path to the file prefixed with an `@`, for example:

```sh
m365 spo sitescript add --title "Contoso" --description "Contoso theme script" --content @themeScript.json
```

CLI for Microsoft 365 will load the contents from the specified file and use it in the command that you specified.

You can use the `@` token in any command, with any option that accepts a value.

## Verbose and debug mode

By default, commands output only the information returned by the corresponding Microsoft 365 API, whether the command result or error. You can choose for a more user-friendly output by using the `--verbose` option or setting the `CLIMICROSOFT365_VERBOSE` environment variable to `1`. For example: by default, when checking status of the Microsoft 365 Public CDN, you would see:

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

If you're experiencing problems when using the CLI for Microsoft 365, you can use the `--debug` option or set the `CLIMICROSOFT365_DEBUG` environment variable to `1`. On top of the output from the verbose mode, the debug mode will provide you with detailed information about all requests and responses from the Microsoft 365 APIs used by the command.

## Command completion

To help you use its commands, CLI for Microsoft 365 offers you the ability to autocomplete commands and options that you're typing in the prompt. Some additional setup, specific for the shell and terminal that you use, is required to enable command completion for CLI for Microsoft 365. For more information on configuring command completion for the CLI for Microsoft 365 see the [command completion](completion.md) article.

## Disable automatic checking for updates

Each time you run CLI for Microsoft 365, it will automatically check if there is a new version available and prompt you with update instructions if that's the case. If you use CLI for Microsoft 365 in CI/CD or in scripts and want to make it run faster, you can disable the check by setting the `CLIMICROSOFT365_NOUPDATE` environment variable to `1`.
