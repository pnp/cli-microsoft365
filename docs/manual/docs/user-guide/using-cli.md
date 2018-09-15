# Using the Office 365 CLI

Information in this section will help you understand how the Office 365 CLI works and how you can use it most effectively.

## Start the CLI

You can use the Office 365 CLI in two modes: immersive, where it launches a separate command prompt, and non-immersive, where you can execute Office 365 CLI commands directly from your shell.

### Start the CLI in the immersive mode

To start Office 365 CLI in the immersive mode, execute in the command line:

```sh
office365
```

or for short:

```sh
o365
```

In both cases, Office 365 CLI will start a new command prompt where you can directly interact with it and its commands.

<script src="https://asciinema.org/a/158205.js" id="asciicast-158205" async></script>

!!! info
    When using the Office 365 CLI in immersive mode, it starts a new command prompt. In this command prompt, you have access only to Office 365 CLI commands. If you try to run any of the system commands, such as `ls`, you will get an error stating that you tried to execute an unknown command. If you need to use system commands or other CLIs while working with Office 365 CLI, consider using the non-immersive mode instead.

The main benefit of using the Office 365 CLI in the immersive mode, is that you have direct access to all commands, without having to prepend them with `o365` and that you can have the CLI complete your input by pressing `TAB` without any additional configuration.

To close the Office 365 CLI started in immersive mode, type `exit` in the CLI prompt.

### Use the CLI in non-immersive mode

Another way to use the Office 365 CLI is by executing specific commands directly from the command line.

<script src="https://asciinema.org/a/158207.js" id="asciicast-158207" async></script>

!!! attention
    When using the Office 365 CLI this way, each CLI command must be prepended with `office365` or `o365` for short. Without this, your shell will not know how the particular command should be executed.

Using the Office 365 CLI directly from the command line is invaluable if you want to write scripts consisting of a number of Office 365 CLI and other commands combined together. Additionally, you keep the access to all system commands and other CLIs available on your computer.

## List available commands

To list commands available with the Office 365 CLI type `help` in the CLI prompt, or `o365 help` directly in your shell.

Commands in the Office 365 CLI are combined into groups. You can list the commands available in the particular group by typing `help <group>`, for example `help spo` to list all commands related to SharePoint Online, or `o365 help spo` directly in your shell.

<script src="https://asciinema.org/a/158209.js" id="asciicast-158209" async></script>

## Get command help

Each command in the Office 365 CLI comes with help describing the command's purpose, available options and any other relevant details as well as related resources. To get the basic help information with command's description and available options, type `help <command>` or `o365 help <command>` in the shell, for example to get help for the `spo cdn get` command, type in the shell `o365 help spo cdn get`.

To get the complete help information including background information, examples and links to related information, use the `--help` option, for example `o365 spo cdn get --help`. This ability is also useful if you've already typed some options and don't want to lose your input but want to access the help, for example: `o365 spo cdn get --type Private --help`.

<script src="https://asciinema.org/a/158212.js" id="asciicast-158212" async></script>

## Required and optional command options

Commands in the Office 365 CLI often contain options that determine what the particular command should do. Command options vary from the URL of the site for which you would like to retrieve more information, to the type of Office 365 CDN that you would like to manage.

Some options are required and necessary for the particular command to execute, while other are optional. When listing available options for the particular command, Office 365 CLI follows the naming convention where required options are wrapped in angle brackets (`< >`) while optional options are wrapped in square brackets (`[ ]`). For example, in the `spo cdn origin add` command, the origin you want to add is required (`-r, --origin <origin>`), while the type of CDN for which this origin should be added is optional (`-t, --type [type]`) and its value defaults to `Public`.

## Values with quotes

In cases, when the option's value contains spaces, it should be wrapped in quotes. For example, to create a modern team site for the _Office 365 CLI_ team, you would execute in the shell:

```sh
o365 spo site add --alias office365cli --title "Office 365 CLI"
```

When the value, that you want to provide contains quotes, it needs to be wrapped in quotes as well, for example to pass a JSON value in the CLI prompt, you would execute:

```sh
spo sitescript add --title "Contoso" --description "Contoso theme script" --content '{"abc": "def"}'
```

If you use the Office 365 CLI in Bash, the outer pair of quotes will be processed by Bash so the value needs to be wrapped in an additional pair of quotes, for example:

```sh
o365 spo sitescript add --title "Contoso" --description "Contoso theme script" --content '`{"abc": "def"}`'
```

## Verbose and debug mode

By default, commands output only the information returned by the corresponding Office 365 API, whether the command result or error. You can choose for a more user-friendly output by using the `--verbose` option or setting the `OFFICE365CLI_VERBOSE` environment variable to `1`. For example: by default, when checking status of the Office 365 Public CDN, you would see:

```sh
$ o365 spo cdn get
true
```

After adding the `--verbose` option, the output would change to:

```sh
$ o365 spo cdn get --verbose
Retrieving status of Public CDN...
Public CDN at https://contoso-admin.sharepoint.com is enabled
```

If you're experiencing problems when using the Office 365 CLI, you can use the `--debug` option or set the `OFFICE365CLI_DEBUG` environment variable to `1`. On top of the output from the verbose mode, the debug mode will provide you with detailed information about all requests and responses from the Office 365 APIs used by the command.

## Command completion

To help you use its commands, the Office 365 CLI offers you the ability to complete commands and options that you're typing in the prompt. Depending how you're using the Office 365 CLI, some additional setup might be required to enable command completion.

### Completion in immersive mode

When using the Office 365 CLI in the immersive mode, the CLI prompt helps you complete your input. By pressing the `TAB` key once, the CLI will complete your current input. By pressing `TAB` twice, it will show you all available commands or options.

<script src="https://asciinema.org/a/158219.js" id="asciicast-158219" async></script>

### Completion in non-immersive mode

Also when running in non-immersive mode, the Office 365 CLI offers you support for completing your input. The configuration steps required to enable command completion, depend on which operating system and shell you're using. For more information on configuring command completion for the Office 365 CLI see the [command completion](../concepts/completion.md) article.