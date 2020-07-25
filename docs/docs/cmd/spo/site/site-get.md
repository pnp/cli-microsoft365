# spo site get

Gets information about the specific site collection

## Usage

```sh
m365 spo site get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: URL of the site collection to retrieve information for

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

This command can retrieve information for both classic and modern sites.

## Examples

Return information about the _https://contoso.sharepoint.com/sites/project-x_ site collection.

```sh
m365 spo site get -u https://contoso.sharepoint.com/sites/project-x
```
