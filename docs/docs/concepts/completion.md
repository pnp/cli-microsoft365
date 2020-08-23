# Command completion

To help you use its commands, CLI for Microsoft 365 offers you the ability to autocomplete commands and options that you're typing in the prompt. Some additional setup, specific for the shell and terminal that you use, is required to enable command completion for CLI for Microsoft 365.

## Clink (cmder)

On Windows, the CLI for Microsoft 365 offers support for completing commands in [cmder](http://cmder.net) and other shells using [Clink](https://mridgers.github.io/clink/).

### Enable Clink completion

To enable completion:

1. Start your shell
1. Change the working directory to where your shell stores completion plugins. For cmder, it's `%CMDER_ROOT%\vendor\clink-completions`, where `%CMDER_ROOT%` is the folder where you installed cmder.
1. Execute: `m365 cli completion clink update > m365.lua`. This will create the `m365.lua` file with information about m365 commands which is used by Clink to provide completion
1. Restart your shell

You should now be able to complete your input, eg. typing `m365 s<tab>` will complete it to `m365 spo` and typing `m365 spo <tab><tab>` will list all SharePoint Online commands available in CLI for Microsoft 365. To see the options available for the current command, type `-<tab><tab>`, for example `m365 spo app list -<tab><tab>` will list all options available for the `m365 spo app list` command.

### Update Clink completion

Command completion is based on a static file. After updating the CLI for Microsoft 365, you should update the completion file as described in the [Enable completion](#enable-clink-completion) section so that the completion file reflects the latest commands in the CLI for Microsoft 365.

### Disable Clink completion

To disable completion, delete the `m365.lua` file you generated previously and restart your shell.

## Zsh, Bash and Fish

If you're using Zsh, Bash or Fish as your shell, you can benefit of CLI for Microsoft 365 command completion as well, when typing commands directly in the shell. The completion is based on the [Omelette](https://www.npmjs.com/package/omelette) package.

For Mac Terminal, you'll need to add `source /usr/local/etc/profile.d/bash_completion.sh` to `~/.bashrc`

### Enable sh completion

To enable completion:

1. Start your shell
1. Execute `m365 cli completion sh setup`. This will generate the `commands.json` file in the same folder where the CLI for Microsoft 365 is installed, listing all available commands and their options. Additionally, it will register completion in your shell profile file (for Zsh `~/.zshrc`) using the [Omelette's automated install](https://www.npmjs.com/package/omelette#automated-install).
1. Restart your shell

You should now be able to complete your input, eg. typing `m365 s<tab>` will complete it to `m365 spo` and typing `m365 spo <tab><tab>` will list all SharePoint Online commands available in CLI for Microsoft 365. To see the options available for the command, type `-<tab><tab>`, for example `m365 spo app list -<tab><tab>` will list all options available for the `m365 spo app list` command. If the command is completed, the completion will automatically start suggestions with a `-` indicating that you have matched a command and can now specify its options. Command options you've already used are removed from the suggestions list, but the completion doesn't take into account short and long variant of the same option. If you specified the `--output` option in your command, `--option` will not be displayed in the list of suggestions, but `-o` will.

### Update sh completion

Command completion is based on the static `commands.json` file located in the folder where the CLI for Microsoft 365 is installed. After updating the CLI for Microsoft 365, you should update the completion file by executing `m365 cli completion sh update` in the command line. After running this command, it's not necessary to restart the shell to see the latest changes.

### Disable sh completion

To disable completion, edit your shell's profile file (for Zsh `~/.zshrc`) and remove the following lines:

```sh
# begin m365 completion
. <(m365 --completion)
# end m365 completion
```

Save the profile file and restart the shell for the changes to take effect.

## PowerShell

If you use CLI for Microsoft 365 in PowerShell you can use the custom argument completer provided with the CLI for Microsoft 365 to get command completion when typing commands directly in the shell.

### Enable PowerShell completion

To enable completion in your current PowerShell profile:

1. Start PowerShell
1. Execute `m365 cli completion pwsh setup --profile $profile`. This will generate the `commands.json` file in the same folder where the CLI for Microsoft 365 is installed, listing all available commands and their options. Additionally, it will register completion in your PowerShell profile
1. Restart PowerShell

You should now be able to complete your input, eg. typing `m365 s<tab>` will complete it to `m365 spo` and typing `m365 spo <tab><tab>` will list all SharePoint Online commands available in CLI for Microsoft 365. To see the options available for the command, type `-<tab><tab>`, for example `m365 spo app list -<tab><tab>` will list all options available for the `m365 spo app list` command. If the command is completed, the completion will automatically start suggestions with a `-` indicating that you have matched a command and can now specify its options. Command options you've already used are removed from the suggestions list, but the completion doesn't take into account short and long variant of the same option. If you specified the `--output` option in your command, `--option` will not be displayed in the list of suggestions, but `-o` will.

### Update PowerShell completion

Command completion is based on the static `commands.json` file located in the folder where the CLI for Microsoft 365 is installed. After updating the CLI for Microsoft 365, you should update the completion file by executing `m365 cli completion pwsh update` in the command line. After running this command, it's not necessary to restart PowerShell to see the latest changes.

### Disable PowerShell completion

To disable CLI for Microsoft 365 command completion in your PowerShell profile, open the profile file in a code editor, and remove the reference to the `Register-O365CLICompletion.ps1` script. Restart PowerShell for the changes to take effect.
