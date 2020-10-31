# aad groupsetting list

Lists Azure AD group settings

## Usage

```sh
m365 aad groupsetting list [options]
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

List all group settings in the tenant

```sh
m365 aad groupsetting list
```