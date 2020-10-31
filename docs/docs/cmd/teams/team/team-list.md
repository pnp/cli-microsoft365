# teams team list

Lists Microsoft Teams teams in the current tenant

## Usage

```sh
m365 teams team list [options]
```

## Options

`-h, --help`
: output usage information

`-j, --joined`
: Show only joined teams

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

You can only see the details or archived status of the Microsoft Teams you are a member of.

## Examples

List all Microsoft Teams in the tenant

```sh
m365 teams team list
```

List all Microsoft Teams in the tenant you are a member of

```sh
m365 teams team list --joined
```