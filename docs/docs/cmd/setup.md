# setup

Sets up CLI for Microsoft 365 based on your preferences

## Usage

```sh
m365 setup [options]
```

## Options

`--interactive`
: Configure CLI for Microsoft 365 for interactive use without prompting for additional information

`--scripting`
: Configure CLI for Microsoft 365 for use in scripts without prompting for additional information

--8<-- "docs/cmd/_global.md"

## Remarks

The `m365 setup` command is a wizard that helps you configure the CLI for Microsoft 365 for your needs. It will ask you a series of questions and based on your answers, it will configure the CLI for Microsoft 365 for you.

The command will ask you the following questions:

- _How do you plan to use the CLI?_

  You can choose between **interactive** and **scripting** use. In interactive mode, the CLI for Microsoft 365 will prompt you for additional information when needed, automatically open links browser, automatically show help on errors and show spinners. In **scripting** mode, the CLI will not use interactivity to prevent blocking your scripts.

- _Are you going to use the CLI in PowerShell?_ (asked only when you chose to configure CLI for scripting)

  To simplify using CLI in PowerShell, you can configure the CLI to output errors to stdout, instead of the default stderr which is tedious to handle in PowerShell.

- _How experienced are you in using the CLI?_

  You can choose between **beginner** and **proficient**. If you're just starting working with the CLI, it will show you full help information. For more experienced users, it will only show information about commands' options. No matter how you configure this setting, you can always invoke full help by using `--help full`.

After you answer these questions, the command will build a preset and ask you to confirm using it to configure the CLI.

If you want to configure the CLI on a non-interactive system, or want to configure CLI without answering the questions, you can use the `--interactive` or `--scripting` options. When using these options, CLI will apply the correct presets. Additionally, it will detect if it's used in PowerShell and configure CLI accordingly.

The `m365 setup` command uses the following presets:

- interactive use:
  - autoOpenLinksInBrowser: true,
  - copyDeviceCodeToClipboard: true,
  - output: 'text',
  - printErrorsAsPlainText: true,
  - prompt: true,
  - showHelpOnFailure: true,
  - showSpinner: true
- scripting use:
  - autoOpenLinksInBrowser: false,
  - copyDeviceCodeToClipboard: false,
  - output: 'json',
  - printErrorsAsPlainText: false,
  - prompt: false,
  - showHelpOnFailure: false,
  - showSpinner: false
- use in PowerShell:
  - errorOutput: 'stdout'
- beginner:
  - helpMode: 'full'
- proficient:
  - helpMode: 'options'

## Examples

Configure CLI for Microsoft based on your preferences interactively

```sh
m365 setup
```

Configure CLI for Microsoft for interactive use without prompting for additional information

```sh
m365 setup --interactive
```

Configure CLI for Microsoft for use in scripts without prompting for additional information

```sh
m365 setup --scripting
```
