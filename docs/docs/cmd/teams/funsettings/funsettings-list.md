# teams funsettings list

Lists fun settings for the specified Microsoft Teams team

## Usage

```sh
m365 teams funsettings list [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the team for which to list fun settings

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

List fun settings of a Microsoft Teams team

```sh
m365 teams funsettings list --teamId 83cece1e-938d-44a1-8b86-918cf6151957
```
