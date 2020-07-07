# aad groupsettingtemplate list

Lists Azure AD group settings templates

## Usage

```sh
aad groupsettingtemplate list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

List all group setting templates in the tenant

```sh
aad groupsettingtemplate list
```