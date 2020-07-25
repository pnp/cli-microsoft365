# spo page control get

Gets information about the specific control on a modern page

## Usage

```sh
m365 spo page control get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: ID of the control to retrieve information for

`-n, --name <name>`
: Name of the page where the control is located

`-u, --webUrl <webUrl>`
: URL of the site where the page to retrieve is located

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

If the specified `name` doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

## Examples

Get information about the control with ID _3ede60d3-dc2c-438b-b5bf-cc40bb2351e1_ placed on a modern page with name _home.aspx_

```sh
m365 spo page control get --id 3ede60d3-dc2c-438b-b5bf-cc40bb2351e1 --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```