# spo web list

Lists subsites of the specified site

## Usage

```sh
m365 spo web list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the parent site for which to retrieve the list of subsites

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Return all subsites from site _https://contoso.sharepoint.com/_

```sh
m365 spo web list -u https://contoso.sharepoint.com
```