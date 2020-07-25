# cli completion sh setup

Sets up command completion for Zsh, Bash and Fish

## Usage

```sh
m365 cli completion sh setup [options]
```

## Options

`-h, --help`
: output usage information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Set up command completion for Zsh, Bash or Fish

```powershell
cli completion sh setup
```

## More information

- Command completion: [https://pnp.github.io/cli-microsoft365/concepts/completion/](https://pnp.github.io/cli-microsoft365/concepts/completion/)
