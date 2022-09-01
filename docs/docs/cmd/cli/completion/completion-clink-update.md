# cli completion clink update

Updates command completion for Clink (cmder)

## Usage

```sh
m365 cli completion clink update [options]
```

## Options

--8<-- "docs/cmd/_global.md"

!!! important
    Before running this command, change the working directory to where your shell stores completion plugins. For cmder, it's `%CMDER_ROOT%\vendor\clink-completions`, where `%CMDER_ROOT%` is the folder where you installed cmder. After running this command, restart your terminal to load the completion.

## Remarks

This commands updates the list of commands and their options used by command completion in Clink (cmder). You should run this command each time after upgrading the CLI for Microsoft 365.

## Examples

Write the list of commands for Clink (cmder) command completion to a file named `m365.lua` in the current directory

```powershell
cli completion clink update > m365.lua
```

## More information

- Command completion: [https://pnp.github.io/cli-microsoft365/user-guide/completion/](https://pnp.github.io/cli-microsoft365/user-guide/completion/)
