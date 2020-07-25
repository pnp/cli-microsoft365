# spo page get

Gets information about the specific modern page

## Usage

```sh
m365 spo page get [options]
```

## Options

`-h, --help`
: output usage information

`-n, --name <name>`
: Name of the page to retrieve

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

If the specified name doesn't refer to an existing modern page, you will get a `File doesn't exists` error.

## Examples

Get information about the modern page with name _home.aspx_

```sh
m365 spo page get --webUrl https://contoso.sharepoint.com/sites/team-a --name home.aspx
```