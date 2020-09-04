# cli completion sh update

Updates command completion for Zsh, Bash and Fish

## Usage

```sh
cli completion sh update [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

This commands updates the list of commands and their options used by command completion in Zsh, Bash and Fish. You should run this command each time after upgrading the CLI for Microsoft 365.

## Examples

Update list of commands for Zsh, Bash and Fish command completion

```sh
cli completion sh update
```

## More information

- Command completion: [https://pnp.github.io/cli-microsoft365/concepts/completion/](https://pnp.github.io/cli-microsoft365/concepts/completion/)
