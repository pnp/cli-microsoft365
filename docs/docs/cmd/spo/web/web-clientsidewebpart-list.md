# spo web clientsidewebpart list

Lists available client-side web parts

## Usage

```sh
m365 spo web clientsidewebpart list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve the information

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Lists all the available client-side web parts for the specified site

```sh
m365 spo web clientsidewebpart list --webUrl https://contoso.sharepoint.com
```