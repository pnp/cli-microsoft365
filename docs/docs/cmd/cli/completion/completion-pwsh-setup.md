# cli completion pwsh setup

Sets up command completion for PowerShell

## Usage

```sh
m365 cli completion pwsh setup [options]
```

## Options

`-h, --help`
: output usage information

`-p, --profile <profile>`
: Path to the PowerShell profile file

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

This commands sets up command completion for the CLI for Microsoft 365 in PowerShell by registering a custom PowerShell argument completer in the specified profile. Because CLI for Microsoft 365 is not a native PowerShell module, it requires a custom completer to provide completion.

If the specified profile path doesn't exist, the CLI will try to create it.

## Examples

Set up command completion for PowerShell using the profile from the profile variable

```powershell
cli completion pwsh setup --profile $profile
```

## More information

- Command completion: [https://pnp.github.io/cli-microsoft365/concepts/completion/](https://pnp.github.io/cli-microsoft365/concepts/completion/)
